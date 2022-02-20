#include "PyRtd.h"
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyEvents.h"
#include "EventLoop.h"
#include <xloil/RtdServer.h>
#include <xloil/ExcelThread.h>
#include <pybind11/pybind11.h>
#include <future>

namespace py = pybind11;
using std::future_status;
using std::shared_ptr;
using std::make_shared;

namespace
{
  // See https://github.com/pybind/pybind11/issues/1389

  template <typename T> class py_shared_ptr {
  private:
    shared_ptr<T> _impl;

  public:
    using element_type = T;

    py_shared_ptr() {}

    py_shared_ptr(T *ptr) 
    {
      auto pyobj = py::cast(ptr);
      auto* pyptr = pyobj.ptr();
      Py_INCREF(pyptr);
      shared_ptr<PyObject> pyObjPtr(
        pyptr, 
        [](PyObject* x) { py::gil_scoped_acquire getGil; Py_XDECREF(x); }
      );
      _impl = shared_ptr<T>(pyObjPtr, ptr);
    }

    py_shared_ptr(std::shared_ptr<T> r) : _impl(r) {}

    operator std::shared_ptr<T>() const { return _impl; }

    T* get() const { return _impl.get(); }
  };
}

PYBIND11_DECLARE_HOLDER_TYPE(T, py_shared_ptr<T>);

namespace xloil
{
  namespace Python
  {
    namespace
    {
      auto& asyncEventLoop()
      {
        return *theCoreAddin().thread;
      }
    }

    RtdReturn::RtdReturn(
      IRtdPublish& notify,
      const shared_ptr<const IPyToExcel>& returnConverter,
      const CallerInfo& caller)
      : _notify(notify)
      , _returnConverter(returnConverter)
      , _caller(caller)
      , _running(true)
    {
    }

    RtdReturn::~RtdReturn()
    {
      if (!_task.ptr())
        return;

      py::gil_scoped_acquire gilAcquired;
      _task = py::object();
    }

    void RtdReturn::set_task(const py::object& task)
    {
      _task = task;
      _running = true;
    }

    void RtdReturn::set_result(const py::object& value) const
    {
      if (!_running)
        return;
      
      // Convert result to ExcelObj
      ExcelObj result = _returnConverter
        ? (*_returnConverter)(*value.ptr())
        : FromPyObj<false>()(value.ptr());

      // If nil, conversion wasn't possible, so use the cache
      if (result.isType(ExcelType::Nil))
        result = pyCacheAdd(value, CallerInfo(_caller));

      _notify.publish(std::move(result));
    }
    void RtdReturn::set_done()
    {
      if (!_running)
        return;
      _running = false;

      _task = py::object();
    }
    void RtdReturn::cancel()
    {
      if (!_running)
        return;
      
      _running = false;
      if (_task)
        asyncEventLoop().callback(_task.attr("cancel"));
    }
    bool RtdReturn::done() noexcept
    {
      return !_running;
    }
    void RtdReturn::wait() noexcept
    {
      // asyncio.Future has no 'wait'
    }
    const CallerInfo& RtdReturn::caller() const noexcept
    {
      return _caller;
    }

    /// <summary>
    /// Uses the RtdReturn object to launch a publishing task.
    /// </summary>
    struct PyRtdTaskLauncher : public IRtdTask
    {
      shared_ptr<RtdReturn> _returnObj;
      CallerInfo _caller;
      PyObjectHolder _func;
      shared_ptr<const IPyToExcel> _returnConverter;

      PyRtdTaskLauncher(const py::object& func, const shared_ptr<IPyToExcel>& converter) 
        : _func(func)
        , _returnConverter(converter)
      {}

      virtual ~PyRtdTaskLauncher()
      {}

      void start(IRtdPublish& publisher) override
      {
        _returnObj.reset(new RtdReturn(publisher, _returnConverter, _caller));
        py::gil_scoped_acquire gilAcquired;
        try
        {
          _func(py::cast(_returnObj));
        }
        catch (const std::exception& e)
        {
          publisher.publish(ExcelObj(e.what()));
        }
      }
      bool done() noexcept override
      {
        return _returnObj ? _returnObj->done() : false;
      }
      void wait() noexcept override
      {
        if (_returnObj)
          _returnObj->wait();
      }
      void cancel() override
      {
        if (_returnObj && !_returnObj->done())
        {
          py::gil_scoped_acquire gilAcquired;
          _returnObj->cancel();
        }
      }
    };

    class PyRtdServer
    {
      shared_ptr<IRtdServer> _impl;
      shared_ptr<const void> _cleanup;
      std::future<shared_ptr<IRtdServer>> _initialiser;

      IRtdServer& impl()
      {
        if (!_impl)
        {
          if (!_initialiser.valid())
            XLO_THROW("RtdServer terminated");
          _impl = _initialiser.get();
        }
        return *_impl;
      }

    public:
      PyRtdServer()
      {
        _initialiser = runExcelThread([]() { return newRtdServer(); });
        // Destroy the Rtd server if we are still around on python exit. The 
        // Rtd server may maintain links to python objects and Excel may not
        // call the server terminate function until after python has unloaded.
        // PyBye will only be called synchronously from the thread destroying the 
        // interpreter, so capturing 'this' is safe.
        _cleanup = Event_PyBye().bind([this]
        { 
          py::gil_scoped_release releaseGil;
          _impl.reset(); 
        });
      }
      ~PyRtdServer()
      {
        py::gil_scoped_release releaseGil;
        _impl.reset();
      }
      void start(const py_shared_ptr<IRtdPublisher>& topic)
      {
        py::gil_scoped_release releaseGil;
        impl().start(topic);
      }
      bool publish(
        const wchar_t* topic, 
        const py::object& value, 
        IPyToExcel* converter=nullptr)
      {
        auto ptr = value.ptr();
        ExcelObj xlValue;

        if (PyExceptionInstance_Check(ptr))
        {
          auto tb = PySteal(PyException_GetTraceback(ptr));

          // We need to set the python error state so that the error_string 
          // function works
          PyErr_Restore(value.get_type().ptr(), value.ptr(), tb.ptr());
          auto errStr = py::detail::error_string();
          // Restore the error state to clear before proceeding to avoid 
          // strange behaviour in the event call.
          PyErr_Clear();

          Event_PyUserException().fire(PyBorrow(value.get_type().ptr()), value, tb);
          xlValue = errStr;
        }
        else
        {
          xlValue = converter
            ? (*converter)(*ptr)
            : FromPyObj()(ptr);
        }

        py::gil_scoped_release releaseGil;
        return impl().publish(topic, std::move(xlValue));
      }
      py::object subscribe(const wchar_t* topic, IPyFromExcel* converter=nullptr)
      {
        shared_ptr<const ExcelObj> value;
        {
          py::gil_scoped_release releaseGil;
          value = impl().subscribe(topic);
        }
        if (!value)
          return py::none();
        return PySteal<>(converter
          ? (*converter)(*value)
          : PyFromAny()(*value));
      }
      py::object peek(const wchar_t* topic, IPyFromExcel* converter = nullptr)
      {
        shared_ptr<const ExcelObj> value;
        {
          py::gil_scoped_release releaseGil;
          value = impl().peek(topic);
        }
        if (!value)
          return py::none();
        return PySteal<>(converter
          ? (*converter)(*value)
          : PyFromAny()(*value));
      }
      void drop(const wchar_t* topic)
      {
        // Don't throw if _impl has been destroyed, just ignore
        if (_impl)
        {
          py::gil_scoped_release releaseGil;
          _impl->drop(topic);
        }
      }

      void startTask(
        const wchar_t* topic, 
        const py::object& func, 
        const shared_ptr<IPyToExcel>& converter = nullptr)
      {
        auto task = make_shared<PyRtdTaskLauncher>(func, converter);

        py::gil_scoped_release releaseGil;
        impl().start(
          make_shared<RtdPublisher>(
            topic, impl(), task));
      }
    };

    class PyRtdTopic : public IRtdPublisher
    {
    public:
      using IRtdPublisher::IRtdPublisher;

      virtual void connect(size_t numSubscribers) override
      {
        PYBIND11_OVERLOAD_PURE(void, IRtdPublisher, connect, numSubscribers)
      }
      virtual bool disconnect(size_t numSubscribers) override
      {
        PYBIND11_OVERLOAD_PURE(bool, IRtdPublisher, disconnect, numSubscribers)
      }
      virtual void stop() override
      {
        PYBIND11_OVERLOAD_PURE(void, IRtdPublisher, stop, )
      }
      virtual bool done() const override
      {
        PYBIND11_OVERLOAD_PURE(bool, IRtdPublisher, done, )
      }
      virtual const wchar_t* topic() const noexcept override
      {
        try
        {
        PYBIND11_OVERLOAD_PURE(const wchar_t *, IRtdPublisher, topic, )
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd publisher failed to get topic name: {}", e.what());
          return L"";
        }
      }
    };
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<IRtdPublisher, PyRtdTopic, py_shared_ptr<IRtdPublisher>>(mod, "RtdPublisher")
          .def(py::init<>())
          .def("connect", &IRtdPublisher::connect)
          .def("disconnect", &IRtdPublisher::disconnect)
          .def("stop", &IRtdPublisher::stop)
          .def("done", &IRtdPublisher::done)
          .def("topic", &IRtdPublisher::topic);

        py::class_<PyRtdServer>(mod, "RtdServer")
          .def(py::init<>())
          .def("start", &PyRtdServer::start,
            py::arg("topic"))
          .def("publish", &PyRtdServer::publish,
            py::arg("topic"), py::arg("value"), py::arg("converter") = nullptr)
          .def("subscribe", &PyRtdServer::subscribe,
            py::arg("topic"), py::arg("converter") = nullptr)
          .def("peek", &PyRtdServer::peek,
            py::arg("topic"), py::arg("converter") = nullptr)
          .def("drop", &PyRtdServer::drop)
          .def("start_task", &PyRtdServer::startTask,
            py::arg("topic"), py::arg("func"), py::arg("converter") = nullptr);

        py::class_<RtdReturn, shared_ptr<RtdReturn>>(mod, "RtdReturn")
          .def("set_result", &RtdReturn::set_result)
          .def("set_done", &RtdReturn::set_done)
          .def("set_task", &RtdReturn::set_task)
          .def_property_readonly("caller", &RtdReturn::caller)
          .def_property_readonly("loop", [](py::object x) { return asyncEventLoop().loop(); });
      });
    }
  }
}