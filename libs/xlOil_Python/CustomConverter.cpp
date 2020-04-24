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
      PyObject* fromString(const wchar_t* buf, size_t len) const
      {
        return PyFromString().fromString(buf, len);
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

    static int theBinder = addBinder([](py::module& mod)
    {
      pybind11::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>
        (mod, "CustomConverter")
        .def(py::init<py::object>(), py::arg("callable"));
    });
  }
}