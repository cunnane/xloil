#include "PyCoreModule.h"
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
        int* toR = nullptr, int* toC = nullptr,
        size_t* nRows = nullptr, size_t* nCols = nullptr)
      {
        py::gil_scoped_release loseGil;
        const auto toRow = toR ? *toR : (nRows ? fromR + (int)*nRows - 1: Range::TO_END);
        const auto toCol = toC ? *toC : (nCols ? fromC + (int)*nCols - 1: Range::TO_END);
        return r.range(fromR, fromC, toRow, toCol);
      }

      // Works like the Range.Cell function in VBA except is zero based
      inline auto rangeCell(const Range& r, int row, int col)
      {
        return r.cell(row, col);
      }
      auto convertExcelObj(ExcelObj&& val)
      {
        return PySteal(PyFromAny()(val));
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
        size_t fromRow, fromCol, toRow, toCol, nRows, nCols;
        std::tie(nRows, nCols) = range.shape();
        bool singleValue = sliceHelper2d(loc, nRows, nCols,
          fromRow, fromCol, toRow, toCol);
        return singleValue
          ? convertExcelObj(range.value((int)fromRow, (int)fromCol))
          : py::cast(range.range((int)fromRow, (int)fromCol, (int)toRow, (int)toCol));
      }

      class PyFromRange : public FromExcelBase<PyObject*>
      {
      public:
        using FromExcelBase::operator();

        PyObject* operator()(RefVal obj) const 
        {
          return pybind11::cast(newXllRange(obj)).release().ptr();
        }
        constexpr wchar_t* failMessage() const { return L"Expected range"; }
      };
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
            py::arg("num_rows") = nullptr,
            py::arg("num_cols") = nullptr)
          .def("cells", rangeCell,
            py::arg("row"), 
            py::arg("col"))
          .def("__getitem__", getItem)
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
              return r.shape();
            });

        bindPyConverter<PyExcelConverter<PyFromRange>>(mod, "Range").def(py::init<>());

      }, 99);
    }
  }
}