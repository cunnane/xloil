#include "BasicTypes.h"
#include <xlOil/ExcelRange.h>
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil 
{
  namespace Python
  {
    PyTypeObject* pyExcelErrorType = nullptr;

    template <class T>
    void declare(pybind11::module& mod, const char* type)
    {
      bindFrom<T>(mod, type).def(py::init<>());
    }
    
    inline auto subrange(const ExcelRange& r, int fromR, int fromC, int* nRows = 0, int* nCols = 0, int* toR = 0, int* toC = 0)
    {
      if (!(toR || nRows))
        XLO_THROW("Must specify end row or number of rows");
      if (!(toC || nCols))
        XLO_THROW("Must specify end column or number of columns");
      return r.range(fromR, fromC, nRows ? *nRows : *toR - fromR + 1, nCols ? *nCols : *toC - fromC + 1);
    }
    static int theBinder = addBinder([](py::module& mod)
    {

     // std::is_base_of< IPyFromExcel, PyFromExcel<PyFromInt>>::value;
      //template <typename Base, typename Derived> using is_strict_base_of = bool_constant<
      //std::is_base_of<Base, Derived>::value && !std::is_same<Base, Derived>::value > ;
      //auto thing = py::detail::is_strict_base_of<IPyFromExcel, PyFromExcel<PyFromInt>>::value;
      py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel");
      py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

      // Bind CellError type to xloil::CellError enum
      auto eType = py::enum_<CellError>(mod, "CellError");
      for (auto e : theCellErrors)
        eType.value(wstring_to_utf8(toWCString(e)).c_str(), e);

      pyExcelErrorType = (PyTypeObject*)eType.get_type().ptr();

      // Bind converters for standard types
      declare<PyFromExcel<PyFromInt>>(mod, "int");
      declare<PyFromExcel<PyFromDouble>>(mod, "float");
      declare<PyFromExcel<PyFromBool>>(mod, "bool");
      declare<PyFromExcel<PyFromString>>(mod, "str");
      declare<PyFromExcel<PyFromAny>>(mod, "object");
      declare<PyFromExcel<PyCacheObject>>(mod, "cached");

 
      // Bind Range class from xloil::ExcelRange
      py::class_<ExcelRange>(mod, "Range")
        .def("range", subrange, py::arg("from_row"), py::arg("from_col"),
          py::arg("num_rows") = -1, py::arg("num_cols") = -1,
          py::arg("to_row")=nullptr, py::arg("to_col")=nullptr
          )
        //.def("range", &ExcelRange::range, py::arg("from_row"), py::arg("from_col"), py::arg("num_rows")=-1, py::arg("num_cols")=-1)
        .def("cell", &ExcelRange::cell, py::arg("row"), py::arg("col"))
        .def_property("value", 
          [](const ExcelRange& r) { return PySteal<>(CheckedFromExcel<PyFromAny>()(r.value())); },
          [](ExcelRange& r, const py::object& value) { r = FromPyObj()(value.ptr()); },
          py::return_value_policy::reference_internal)
        .def("set", [](ExcelRange& r, const py::object& value) { r = FromPyObj()(value.ptr()); })
        .def("clear", &ExcelRange::clear)
        .def("address", &ExcelRange::address, py::arg("local")=false)
        .def_property_readonly("num_rows", &ExcelRange::nRows)
        .def_property_readonly("num_cols", &ExcelRange::nCols);
    });

    PyObject * PyFromAny::fromError(CellError err) const
    {
      auto pyObj = py::cast(err);
      return pyObj.release().ptr();
    }
    PyObject * PyFromAny::fromRef(const ExcelObj & obj) const
    {
      return py::cast(ExcelRange(obj)).release().ptr();
    }
  }
}
