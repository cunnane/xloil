#include "PyHelpers.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "InjectedModule.h"
#include "PyExcelArray.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    class UserImpl : public PyFromAny<UserImpl>
    {
    public:
      PyObject* fromArray(const ExcelObj& obj) const
      {
        return py::cast(PyExcelArray(ExcelArray(obj))).release().ptr();
      }
      PyObject* fromArrayObj(const ExcelArray& arr) const
      {
        return py::cast(PyExcelArray(arr)).release().ptr();
      }
      // Override cache ref lookup?
      PyObject* fromString(const PStringView<>& pstr) const
      {
        return PyFromString().fromString(pstr);
      }
    };

    class CustomConverter : public IPyFromExcel
    {
    private:
      py::object _callable;
    public:
      CustomConverter(py::object&& callable)
        : _callable(callable)
      {}
      virtual result_type operator()(const ExcelObj& xl, const_result_ptr defaultVal) const
      {
        auto arg = PySteal<>(PyFromExcel<UserImpl>()(xl, defaultVal));
        return _callable(arg).release().ptr();
      }
      virtual PyObject* fromArray(const ExcelArray& arr) const
      {
        auto pyArr = PyExcelArray(arr);
        auto arg = py::cast(PyExcelArray(arr));
        auto ret = _callable(arg);
        if (pyArr.refCount() > 1)
          XLO_THROW("You cannot keep references to ExcelArray objects");
        return ret.release().ptr();
      }
    };

    py::object cannotConvertException;

    class CustomReturn : public IPyToExcel
    {
    private:
      py::object _callable;
    public:
      CustomReturn(py::object&& callable)
        : _callable(callable)
      {}
      virtual ExcelObj operator()(const PyObject& pyObj) const
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
        return FromPyObj()(converted.ptr(), false);
      }
    };

    class CannotConvert {};

    static int theBinder = addBinder([](py::module& mod)
    {
      py::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>
        (mod, "CustomConverter")
        .def(py::init<py::object>(), py::arg("callable"));

      py::class_<CustomReturn, IPyToExcel, std::shared_ptr<CustomReturn>>
        (mod, "CustomReturn")
        .def(py::init<py::object>(), py::arg("callable"));

      cannotConvertException = py::exception<CannotConvert>(mod, "CannotConvert");
    });
  }
}