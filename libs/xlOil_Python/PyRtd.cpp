#include "PyCoreModule.h"
#include "BasicTypes.h"
#include <xloil/RtdServer.h>
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


    /*
    lass coro:
    def __init__(self):
        self._state = 0

    def __iter__(self):
        return self

    def __await__(self):
        return self

    def __next__(self):
        if self._state == 0:
            self._x = foo()
            self._bar_iter = bar().__await__()
            self._state = 1

        if self._state == 1:
            try:
                suspend_val = next(self._bar_iter)
                # propagate the suspended value to the caller
                # don't change _state, we will return here for
                # as long as bar() keeps suspending
                return suspend_val
            except StopIteration as stop:
                # we got our value
                y = stop.value
            # since we got the value, immediately proceed to
            # invoking `baz`
            baz(self._x, y)
            self._state = 2
            # tell the caller that we're done and inform
            # it of the return value
            raise StopIteration(42)

        # the final state only serves to disable accidental
        # resumption of a finished coroutine
        raise RuntimeError("cannot reuse already awaited coroutine")
    */

    template<class T>
    class PyFuture
    {
      std::future<T> _future;
    public:
      PyFuture(std::future<T> future)
        : _future(future)
      {}
      T await()
      {
        if (_future.wait_for(microseconds(0)) == future_status::ready)
          return this;
          //throw pybind11::stop_iteration()
        return
      }
    };

    class PyRtdServer
    {
      shared_ptr<IRtdServer> _impl;
      shared_ptr<const void> _cleanup;

    public:
      PyRtdServer()
      {
        _impl = newRtdServer();
        // Destroy the Rtd server if we are still around on python exit. The 
        // Rtd server may maintain links to python objects and Excel may not
        // call the server terminate function until after python has unloaded.
        _cleanup = Event_PyBye().bind([self = this] 
        { 
          self->_impl.reset(); 
        });
      }
      ~PyRtdServer()
      {}
      void start(const py_shared_ptr<IRtdPublisher>& topic)
      {
        impl().start(topic);
      }
      bool publish(
        const wchar_t* topic, 
        const py::object& value, 
        IPyToExcel* converter=nullptr)
      {
        return impl().publish(topic, converter
            ? (*converter)(*value.ptr()) 
            : FromPyObj()(value.ptr()));
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
        if (_impl)
          _impl->drop(topic);
      }

      IRtdServer& impl() const
      {
        if (!_impl)
          XLO_THROW("RtdServer terminated");
        return *_impl;
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
      virtual const wchar_t * topic() const override
      {
        PYBIND11_OVERLOAD_PURE(const wchar_t *, IRtdPublisher, topic, )
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