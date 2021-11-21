#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyEvents.h"
#include <xloil/RtdServer.h>
#include <xloil/ExcelApp.h>
#include <pybind11/pybind11.h>
#include <future>
#include <chrono>

namespace py = pybind11;
using std::chrono::microseconds;
using std::future_status;
using std::shared_ptr;

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
        pyptr, [](PyObject *x) { Py_DECREF(x); });
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
          _impl.reset(); 
        });
      }

      void start(const py_shared_ptr<IRtdPublisher>& topic)
      {
        impl().start(topic);
      }
      bool publish(
        const wchar_t* topic, 
        const py::object& value, 
        IPyToExcel* converter=nullptr)
      {
        auto ptr = value.ptr();
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
          return impl().publish(topic, ExcelObj(errStr));
        }
        return impl().publish(topic, converter
            ? (*converter)(*ptr)
            : FromPyObj()(ptr));
      }
      py::object subscribe(const wchar_t* topic, IPyFromExcel* converter=nullptr)
      {
        auto value = impl().subscribe(topic);
        if (!value)
          return py::none();
        return PySteal<>(converter
          ? (*converter)(*value)
          : PyFromAny()(*value));
      }
      py::object peek(const wchar_t* topic, IPyFromExcel* converter = nullptr)
      {
        auto value = impl().peek(topic);
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
          _impl->drop(topic);
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
      virtual const wchar_t * topic() const noexcept override
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
          .def("drop", &PyRtdServer::drop);
      });
    }
  }
}