#include "InjectedModule.h"
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

    class PyRtdManager
    {
      shared_ptr<IRtdManager> _impl;
      shared_ptr<const void> _cleanup;

    public:
      PyRtdManager()
      {
        _impl = newRtdManager();
        // Destroy the Rtd server if we are still around on python exit. The 
        // Rtd server may maintain links to python objects and Excel may not
        // call the server terminate function until after python has unloaded.
        _cleanup = Event_PyBye().bind([self = this] 
        { 
          self->_impl.reset(); 
        });
      }
      ~PyRtdManager()
      {}
      void start(const py_shared_ptr<IRtdTopic>& topic)
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

      IRtdManager& impl() const
      {
        if (!_impl)
          XLO_THROW("RtdManager terminated");
        return *_impl;
      }
    };

    class PyRtdTopic : public IRtdTopic
    {
    public:
      using IRtdTopic::IRtdTopic;

      virtual void connect(size_t numSubscribers) override
      {
        PYBIND11_OVERLOAD_PURE(void, IRtdTopic, connect, numSubscribers)
      }
      virtual bool disconnect(size_t numSubscribers) override
      {
        PYBIND11_OVERLOAD_PURE(bool, IRtdTopic, disconnect, numSubscribers)
      }
      virtual void stop() override
      {
        PYBIND11_OVERLOAD_PURE(void, IRtdTopic, stop, )
      }
      virtual bool done() const override
      {
        PYBIND11_OVERLOAD_PURE(bool, IRtdTopic, done, )
      }
      virtual const wchar_t * topic() const override
      {
        PYBIND11_OVERLOAD_PURE(const wchar_t *, IRtdTopic, topic, )
      }
    };
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<IRtdTopic, PyRtdTopic, py_shared_ptr<IRtdTopic>>(mod, "RtdTopic")
          .def(py::init<>())
          .def("connect", &IRtdTopic::connect)
          .def("disconnect", &IRtdTopic::disconnect)
          .def("stop", &IRtdTopic::stop)
          .def("done", &IRtdTopic::done)
          .def("topic", &IRtdTopic::topic);

        py::class_<PyRtdManager>(mod, "RtdManager")
          .def(py::init<>())
          .def("start", &PyRtdManager::start,
            py::arg("topic"))
          .def("publish", &PyRtdManager::publish,
            py::arg("topic"), py::arg("value"), py::arg("converter") = nullptr)
          .def("subscribe", &PyRtdManager::subscribe,
            py::arg("topic"), py::arg("converter") = nullptr)
          .def("peek", &PyRtdManager::peek,
            py::arg("topic"), py::arg("converter") = nullptr)
          .def("drop", &PyRtdManager::drop);
      });
    }
  }
}