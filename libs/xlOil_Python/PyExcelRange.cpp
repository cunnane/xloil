#include "InjectedModule.h"
#include "PyHelpers.h"
#include "BasicTypes.h"
#include <xlOil/ExcelRange.h>

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      // Works like the Range.Range function in VBA except is zero-based
      inline auto subRange(const Range& r,
        int fromR, int fromC,
        int* toR = 0, int* toC = 0, 
        size_t* nRows = 0, size_t* nCols = 0)
      {
        py::gil_scoped_release loseGil;
        if (!(toR || nRows))
          XLO_THROW("Must specify end row or number of rows");
        if (!(toC || nCols))
          XLO_THROW("Must specify end column or number of columns");
        return r.range(fromR, fromC,
          toR ? *toR : fromR + (int)*nRows,
          toC ? *toC : fromC + (int)*nCols);
      }

      // Works like the Range.Cell function in VBA which is 1-based
      inline auto rangeCell(const Range& r, int row, int col)
      {
        return r.cell(row - 1, col - 1);
      }
      auto convertExcelObj(ExcelObj&& val)
      {
        return PySteal<>(PyFromExcel<PyFromAny<>>()(val));
      }
      auto rangeGetValue(const Range& r)
      {
        return convertExcelObj(r.value());
      }
      void rangeSetValue(Range& r, const py::object& pyval)
      {
        const auto val = FromPyObj()(pyval.ptr());
        // Release gil when setting values in as this may trigger calcs 
        // which call back into other python functions.
        py::gil_scoped_release loseGil;
        r = val;
      };

      void rangeClear(Range& r)
      {
        // Release gil - see reasons above
        py::gil_scoped_release loseGil;
        r.clear();
      }

      py::object getItem(const Range& range, pybind11::tuple loc)
      {
        size_t fromRow, fromCol, toRow, toCol;
        bool singleValue = sliceHelper2d(loc, range.nRows(), range.nCols(), 
          fromRow, fromCol, toRow, toCol);
        return singleValue
          ? convertExcelObj(range.value((int)fromRow, (int)fromCol))
          : py::cast(range.range(fromRow, fromCol, toRow, toCol));
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind Range class from xloil::ExcelRange
        auto rType = py::class_<Range>(mod, "Range")
          .def(py::init([](const wchar_t* x) { return newRange(x); }),
            py::arg("address"))
          .def("range", subRange,
            py::arg("from_row"), 
            py::arg("from_col"),
            py::arg("to_row") = nullptr,
            py::arg("to_col") = nullptr,
            py::arg("num_rows") = -1, 
            py::arg("num_cols") = -1)
          .def("cells", rangeCell,
            py::arg("row"), 
            py::arg("col"))
          .def_property("value",
            rangeGetValue,
            rangeSetValue,
            py::return_value_policy::automatic)
          .def("set", rangeSetValue)
          .def("clear", rangeClear)
          .def("address", [](const Range& r, bool local) { return r.address(local); },
            py::arg("local") = false)
          .def_property_readonly("nrows", [](const Range& r) { return r.nRows(); })
          .def_property_readonly("ncols", [](const Range& r) { return r.nCols(); })
          .def_property_readonly("shape", 
            [](const Range& r)
            {
              return std::make_pair(r.nRows(), r.nCols());
            });

      }, 99);
    }
  }
}