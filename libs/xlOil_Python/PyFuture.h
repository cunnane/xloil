
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
        //Iter(std::future<ExcelObj>&& future)
        //  : _future(std::move(future))
        //{}
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
              throw py::value_error();

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
          value = _iter._future.get();
        }
        return PySteal(TConverter()(value));
      }

      Iter& await() { return _iter; }

      static void bind(pybind11::module& mod, const std::string& name)
      {
        py::class_<Iter>(mod, (name + "Iter").c_str())
          .def("__next__", &Iter::next)
          .def("__iter__", [](py::object self) { return self; });

        py::class_<this_type>(mod, name.c_str())
          .def("__await__", &await, py::return_value_policy::reference_internal);
      }
    };
  }
}
