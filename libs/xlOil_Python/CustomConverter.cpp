#include "PyHelpers.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "PyCoreModule.h"
#include "PyExcelArray.h"
#include "PyEvents.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    class UserImpl : public detail::PyFromAny
    {
    public:
      using detail::PyFromAny::operator();

      PyObject* operator()(ArrayVal arr) const
      {
        return py::cast(PyExcelArray(arr)).release().ptr();
      }
      // Override cache ref lookup?
      PyObject* operator()(const PStringView<>& pstr) const
      {
        return detail::PyFromString()(pstr);
      }
      constexpr wchar_t* failMessage() const { return L"Custom converter failed"; }
    };

    class CustomConverter : public IPyFromExcel
    {
    private:
      py::object _callable;
    public:
      CustomConverter(py::object&& callable)
        : _callable(callable)
      {}
      virtual ~CustomConverter()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }
      virtual result_type operator()(const ExcelObj& xl, const_result_ptr defaultVal) const override
      {
        try
        {
          auto arg = PySteal(PyFromExcel<UserImpl>()(xl, defaultVal));
          return _callable(arg).release().ptr();
        }
        catch (const py::error_already_set& e)
        {
          Event_PyUserException().fire(e.type(), e.value(), e.trace());
          throw;
        }
      }
    };

    py::object cannotConvertException;

    namespace
    {
      auto cleanupConvertException = Event_PyBye().bind([] {
        cannotConvertException = py::object();
      });
    }

    class CustomReturn : public IPyToExcel
    {
    private:
      py::object _callable;
    public:
      CustomReturn(py::object&& callable)
        : _callable(callable)
      {}
      virtual ~CustomReturn()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }
      virtual ExcelObj operator()(const PyObject& pyObj) const override
      {
        // Use raw C API for extra speed as this code is on a critical path
        auto p = (PyObject*)&pyObj;
        auto args = PyTuple_New(1);
        PyTuple_SET_ITEM(args, 0, p);
        Py_INCREF(p); // SetItem steals a reference
        auto result = PyObject_CallObject(_callable.ptr(), args);
        Py_DECREF(args);
        if (!result)
        {
          auto error = PyErr_Occurred();
          if (!PyErr_GivenExceptionMatches(error, cannotConvertException.ptr()))
            throw py::error_already_set();
          PyErr_Clear();
          return ExcelObj();
        }
        auto converted = PySteal<>(result);
        // TODO: the custom converter should be able to specify the return type rather than generic FromPyObj
        // TODO: the user could create an infinite loop which cycles between two type converters - best way to avoid?
        return FromPyObj()(converted.ptr());
      }

      const py::object& handler() const { return _callable; }
    };

    class CannotConvert {};

    static int theBinder = addBinder([](py::module& mod)
    {
      py::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>
        (mod, "CustomConverter")
        .def(py::init<py::object>(), py::arg("callable"));

      py::class_<CustomReturn, IPyToExcel, std::shared_ptr<CustomReturn>>
        (mod, "CustomReturn")
        .def(py::init<py::object>(), py::arg("callable"))
        .def("get_handler", &CustomReturn::handler);

      cannotConvertException = py::exception<CannotConvert>(mod, "CannotConvert");
    });
  }
}