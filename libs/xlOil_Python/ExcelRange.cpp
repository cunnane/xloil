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
      // Works like the Range.Range function in VBA which is 1-based and
      // includes the right hand end-point
      inline auto subRange(const ExcelRange& r,
        int fromR, int fromC,
        int* toR = 0, int* toC = 0, 
        size_t* nRows = 0, size_t* nCols = 0)
      {
        if (!(toR || nRows))
          XLO_THROW("Must specify end row or number of rows");
        if (!(toC || nCols))
          XLO_THROW("Must specify end column or number of columns");
        --fromR; --fromC; // Correct for 1-based indexing
        return r.range(fromR, fromC,
          toR ? *toR : fromR + *nRows,
          toC ? *toC : fromC + *nCols);
      }

      // Works like the Range.Cell function in VBA which is 1-based
      inline auto rangeCell(const ExcelRange& r, int row, int col)
      {
        return r.cells(row - 1, col - 1);
      }

      auto createRange(const wchar_t* address,
        std::optional<int> fromR, std::optional<int> fromC,
        int* nRows = 0, int* nCols = 0,
        int* toR = 0, int* toC = 0)
      {

      }

      auto rangeGetValue(const ExcelRange& range)
      {
        return PySteal<>(PyFromExcel<PyFromAny<>>()(range.value()));
      }
      void rangeSetValue(ExcelRange& r, const py::object& value)
      {
        r = FromPyObj()(value.ptr());
      };

      py::object getItem(const ExcelRange& range, pybind11::tuple loc)
      {
        if (loc.size() != 2)
          XLO_THROW("Expecting tuple of size 2");
        auto r = loc[0];
        auto c = loc[1];
        size_t fromRow, fromCol, toRow, toCol, step = 1;
        if (r.is_none())
        {
          fromRow = 0;
          toRow = range.nRows();
        }
        else if (PySlice_Check(r.ptr()))
        {
          size_t sliceLength;
          r.cast<py::slice>().compute(range.nRows(), &fromRow, &toRow, &step, &sliceLength);
        }
        else
        {
          fromRow = r.cast<size_t>();
          toRow = fromRow + 1;
        }

        if (r.is_none())
        {
          fromCol = 0;
          toCol = range.nRows();
        }
        else if (PySlice_Check(c.ptr()))
        {
          size_t sliceLength;
          c.cast<py::slice>().compute(range.nCols(), &fromCol, &toCol, &step, &sliceLength);
        }
        else
        {
          fromCol = c.cast<size_t>();
          // Check for single element access
          if (fromRow == toRow + 1)
            return rangeGetValue(range.cells((int)fromRow, (int)fromCol));
          toCol = fromCol + 1;
        }

        if (step != 1)
          XLO_THROW("Slices step size must be 1");
        
        return py::cast<ExcelRange>(range.range(fromRow, fromCol, toRow, toCol));
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind Range class from xloil::ExcelRange
        auto rType = py::class_<ExcelRange>(mod, "Range")
          .def(py::init<const wchar_t*>(), 
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
          .def("clear", &ExcelRange::clear)
          .def("address", &ExcelRange::address, 
            py::arg("local") = false)
          .def_property_readonly("nrows", &ExcelRange::nRows)
          .def_property_readonly("ncols", &ExcelRange::nCols)
          .def_property_readonly("shape", 
            [](const ExcelRange& r)
            {
              return std::make_pair(r.nRows(), r.nCols());
            });

      }, 99);
    }
  }
}