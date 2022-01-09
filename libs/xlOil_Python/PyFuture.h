
#pragma once
#include "PyHelpers.h"
#include <future>

namespace xloil
{
  namespace Python
  {
    namespace detail
    {
      class StopIteration : public pybind11::builtin_exception
      {
        PyObject* _value;
      public:
        using builtin_exception::builtin_exception;
        StopIteration(PyObject* value) : StopIteration("") { _value = value; }
        void set_error() const override { PyErr_SetObject(PyExc_StopIteration, _value); }
        // Is PyErr_SetObject stealing the ref?  If not, need to add a decref; 
      };

      struct CastFutureConverter
      {
        template<class T> auto operator()(T x) { return py::cast(x).release().ptr(); }
        template<>        auto operator()(PyObject* x) { return PySteal(x); }
      };
    }
    /// <summary>
    /// This is not a fully flexible wrapper for a std::future.  At the time of writing (Nov 2021)
    /// there is an active pybind11 PR to add async capabilities.
    /// </summary>
    /// <typeparam name="TValType"></typeparam>
    /// <typeparam name="TConverter"></typeparam>
    template <class TValType, class TConverter = detail::CastFutureConverter>
    class PyFuture
    {
      // We follow the recipe in the following link to get a python async object
      // https://stackoverflow.com/questions/51029111/python-how-to-implement-a-c-function-as-awaitable-coroutine
      // 
      //   * Define __await__ to return an iterator. This could be *self* but we want to
      //     avoid the PyFuture object being recognised as an iterator by return converters
      //   * Define the iterator's __iter__ to return self
      //   * Define the iterator's __next__ to return None until the future is ready, then to 
      //     raise StopIteration, passing the result value
      //

    public:
      using this_type = PyFuture<TValType, TConverter>;

      struct Iter
      {
        std::future<TValType> _future;

        /// <summary>
        /// Return None until the future is ready, then raises StopIteration, passing the result value 
        /// </summary>
        void next()
        {
          TValType value;
          {
            pybind11::gil_scoped_release releaseGil;

            if (!_future.valid())
              throw pybind11::value_error();

            if (!_future._Is_ready())
              return;
            value = _future.get();
          }
          throw detail::StopIteration(TConverter()(value));
        };
      };

      Iter _iter;

      /// <summary>
      /// Consumes a std::future<T> to give a PyFuture<T>
      /// </summary>
      /// <param name="future"></param>
      PyFuture(std::future<TValType>&& future)
        : _iter{ std::move(future) }
      {}

      /// <summary>
      /// Synchronously gets the result from the future. Blocking.
      /// </summary>
      /// <returns></returns>
      pybind11::object result()
      {
        TValType value;
        {
          pybind11::gil_scoped_release releaseGil;
          _iter._future.wait();
          value = _iter._future.get();
        }
        return PySteal(TConverter()(value));
      }

      bool done()
      {
        return !_iter._future.valid() || _iter._future._Is_ready();
      }

      Iter& await() { return _iter; }

      static void bind(pybind11::module& mod, const std::string& name)
      {
        pybind11::class_<Iter>(mod, (name + "Iter").c_str())
          .def("__next__", &Iter::next)
          .def("__iter__", [](pybind11::object self) { return self; });

        pybind11::class_<this_type>(mod, name.c_str())
          .def("__await__", &await, pybind11::return_value_policy::reference_internal)
          .def("result", &result)
          .def("done", &done);
      }
    };

    template <>
    class PyFuture<void, void>
    {
    public:
      using this_type = PyFuture<void, void>;

      struct Iter
      {
        std::future<void> _future;

        /// <summary>
        /// Return None until the future is ready, then raises StopIteration, passing the result value 
        /// </summary>
        void next()
        {
          {
            pybind11::gil_scoped_release releaseGil;

            if (!_future.valid())
              throw pybind11::value_error();

            if (!_future._Is_ready())
              return;
            _future.get();
          }
          throw detail::StopIteration(Py_None);
        };
      };

      Iter _iter;

      /// <summary>
      /// Consumes a std::future<T> to give a PyFuture<T>
      /// </summary>
      /// <param name="future"></param>
      PyFuture(std::future<void>&& future)
        : _iter{ std::move(future) }
      {}

      /// <summary>
      /// Synchronously gets the result from the future. Blocking.
      /// </summary>
      /// <returns></returns>
      pybind11::object result()
      {
        {
          pybind11::gil_scoped_release releaseGil;
          _iter._future.get();
        }
        return pybind11::none();
      }

      bool done()
      {
        return !_iter._future.valid() || _iter._future._Is_ready();
      }

      Iter& await() { return _iter; }

      static void bind(pybind11::module& mod)
      {
        pybind11::class_<Iter>(mod, "_FutureIter")
          .def("__next__", &Iter::next)
          .def("__iter__", [](pybind11::object self) { return self; });

        pybind11::class_<this_type>(mod, "_Future")
          .def("__await__", &await, pybind11::return_value_policy::reference_internal)
          .def("result", &result)
          .def("done", &done);
      }
    };
  }
}
