#include "PyRtd.h"
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyEvents.h"
#include "EventLoop.h"
#include "PyAddin.h"
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
        return *theCoreAddin()->thread;
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
      
      XLO_TRACE(L"Received result for RTD task started in '{0}'", _caller.address());

      try
      {
        // Convert result to ExcelObj
        ExcelObj result = _returnConverter
          ? (*_returnConverter)(*value.ptr())
          : FromPyObjOrError()(value.ptr());

        // If nil, conversion wasn't possible, so use the cache
        if (result.isType(ExcelType::Nil))
          result = pyCacheAdd(value, _caller.address().c_str());

        _notify.publish(std::move(result));
      }
      catch (const std::exception& e)
      {
        _notify.publish(ExcelObj(e.what()));
      }
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

      XLO_TRACE(L"Sending cancellation to RTD task started in '{0}'", _caller.address());
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

      /// <summary>
      /// This curious bit of code is designed to execute _debugpy_activate_thread
      /// in the RTD worker thread.  That thread is inaccesible unless you are an
      /// IRtdPublisher
      /// </summary>
      struct InitDebugPy : public IRtdPublisher
      {
        InitDebugPy(IRtdServer* server) : _server(server) {}
        IRtdServer* _server;

        virtual void connect(size_t /*numSubscribers*/)
        {
          if (!_server)
            return;
          py::gil_scoped_acquire getGil;
          pybind11::module::import("xloil.debug").attr("_debugpy_activate_thread")();
          _server->testDisconnect(topicId);
          _server = nullptr;
        }
        virtual bool disconnect(size_t) { return true; }
        virtual void stop() {}
        virtual bool done() const { return true; }
        virtual const wchar_t* topic() const noexcept { return topicName; }

        static constexpr wchar_t* topicName = L"_xlOil_Rtd_Init_";
        static constexpr long topicId = -6666666;
      };
    public:
      PyRtdServer()
      {
        // We don't need the COM or XLL APIs so flags = 0
        _initialiser = runExcelThread([]() 
        { 
          auto server = newRtdServer();
          // InitDebugPy is disabled because as of Aug 2022, VS Code refuses to hit
          // breakpoints in async / rtd functions. VS Pro manages fine, so it's not
          // clear the problem is with xlOil.
          // server->start(make_shared<InitDebugPy>(server.get()));
          // server->testConnect(InitDebugPy::topicId, InitDebugPy::topicName);
          return server;
        }, 0);
        
        // Destroy the Rtd server if we are still around on python exit. The 
        // Rtd server may maintain links to python objects and Excel may not
        // call the server terminate function until after python has unloaded.
        // PyBye will only be called synchronously from the thread destroying the 
        // interpreter, so capturing 'this' is safe.
        _cleanup = Event_PyBye().bind([this]
        { 
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
      
      auto progId() { return impl().progId(); }
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
        py::class_<IRtdPublisher, PyRtdTopic, py_shared_ptr<IRtdPublisher>>(mod, "RtdPublisher",
          R"(
            RTD servers use a publisher/subscriber model with the 'topic' as the key
            The publisher class is linked to a single topic string.

            Typically the publisher will do nothing on construction, but when it detects
            a subscriber using the connect() method, it creates a background publishing task
            When disconnect() indicates there are no subscribers, it cancels this task with
            a call to stop()

            If the task is slow to return or spin up, it could be started the constructor  
            and kept it running permanently, regardless of subscribers.

            The publisher should call RtdServer.publish() to push values to subscribers.
          )")
          .def(py::init<>(), R"(
            This __init__ method must be called explicitly by subclasses or pybind
            will fatally crash Excel.
          )")
          .def("connect", 
            &IRtdPublisher::connect,
            R"(
              Called by the RtdServer when a sheet function subscribes to this 
              topic. Typically a topic will start up its publisher on the first
              subscriber, i.e. when num_subscribers == 1
            )",
            py::arg("num_subscribers"))
          .def("disconnect", 
            &IRtdPublisher::disconnect,
            R"(
              Called by the RtdServer when a sheet function disconnects from this 
              topic. This happens when the function arguments are changed the
              function deleted. Typically a topic will shutdown its publisher 
              when num_subscribers == 0.

              Whilst the topic remains live, it may still receive new connection
              requests, so generally avoid finalising in this method.
            )",
            py::arg("num_subscribers"))
          .def("stop", 
            &IRtdPublisher::stop, 
            R"(
              Called by the RtdServer to indicate that a topic should shutdown
              and dependent threads or tasks and finalise resource usage
            )")
          .def("done", 
            &IRtdPublisher::done,
            R"(
              Returns True if the topic can safely be deleted without leaking resources.
            )")
          .def_property_readonly("topic", 
            &IRtdPublisher::topic, 
            "Returns the name of the topic");

        py::class_<PyRtdServer>(mod, "RtdServer",
          R"(
            An RtdServer allows asynchronous interaction with Excel by posting update
            notifications which cause Excel to recalcate certain cells.  The python 
            RtdServer object manages an RTD COM server with each new RtdServer creating
            an underlying COM server. The RtdServer works on a publisher/subscriber
            model with topics identified by a string. 

            A topic publisher is registered using start(). Subsequent calls to subscribe()
            will connect this topic and tell Excel that the current calling cell should be
            recalculated when a new value is published.

            RTD sits outside of Excel's normal calc cycle: publishers can publish new values 
            at any time, triggering a re-calc of any cells containing subscribers. Note the
            re-calc will only happen 'live' if Excel's caclulation mode is set to automatic
          )")
          .def(py::init<>())
          .def("start", 
            &PyRtdServer::start,
            R"(
              Registers an RtdPublisher with this manager. The RtdPublisher receives
              notification when the number of subscribers changes
            )",
            py::arg("topic"))
          .def("publish", 
            &PyRtdServer::publish,
            R"(
              Publishes a new value for the specified topic and updates all subscribers.
              This function can be called even if no RtdPublisher has been started.

              This function does not use any Excel API and is safe to call at any time
              on any thread.

              An Exception object can be passed at the value, this will trigger the debugging
              hook if it is set. The exception string and it's traceback will be published.
            )",
            py::arg("topic"), 
            py::arg("value"), 
            py::arg("converter") = nullptr)
          .def("subscribe", 
            &PyRtdServer::subscribe,
            R"(
              Subscribes to the specified topic. If no publisher for the topic currently 
              exists, it returns None, but the subscription is held open and will connect
              to a publisher created later. If there is no published value, it will return 
              CellError.NA.  
        
              This calls Excel's RTD function, which means the calling cell will be
              recalculated every time a new value is published.

              Calling this function outside of a worksheet function called by Excel may
              produce undesired results and possibly crash Excel.
            )",
            py::arg("topic"), 
            py::arg("converter") = nullptr)
          .def("peek", 
            &PyRtdServer::peek,
            R"(
              Looks up a value for a specified topic, but does not subscribe.
              If there is no active publisher for the topic, it returns None.
              If there is no published value, it will return CellError.NA.

              This function does not use any Excel API and is safe to call at
              any time on any thread.
            )",
            py::arg("topic"), 
            py::arg("converter") = nullptr)
          .def("drop", 
            &PyRtdServer::drop,
            R"(
              Drops the producer for a topic by calling `RtdPublisher.stop()`, then waits
              for it to complete and publishes #N/A to all subscribers.
            )")
          .def("start_task", 
            &PyRtdServer::startTask,
            R"(
              Launch a publishing task for a `topic` given a func and a return converter.
              The function should take a single `xloil.RtdReturn` argument.
            )",
            py::arg("topic"), 
            py::arg("func"), 
            py::arg("converter") = nullptr)
          .def_property_readonly("progid", &PyRtdServer::progId);

        py::class_<RtdReturn, shared_ptr<RtdReturn>>(mod, "RtdReturn")
          .def("set_result", 
            &RtdReturn::set_result)
          .def("set_done", 
            &RtdReturn::set_done,
            R"(
              Indicates that the task has completed and the RtdReturn can drop its reference
              to the task. Further calls to `set_result()` will be ignored.
            )")
          .def("set_task", 
            &RtdReturn::set_task,
            R"(
              Set the task object to keep it alive until the task indicates it is done. The
              task object should respond to the `cancel()` method.
            )",
            py::arg("task"))
          .def_property_readonly("caller", &RtdReturn::caller)
          .def_property_readonly("loop", [](py::object x) { return asyncEventLoop().loop(); });
      });
    }
  }
}