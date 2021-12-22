#include "PyCore.h"
#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCOM.h"
#include <xlOil/ExcelRange.h>
#include <xlOil/AppObjects.h>

using std::shared_ptr;
using std::wstring_view;
using std::vector;
using std::wstring;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      // Works like the Range.Range function in VBA except is zero-based
      template <class T>
      inline auto subRange(const T& r,
        int fromR, int fromC,
        int* toR = nullptr, int* toC = nullptr,
        size_t* nRows = nullptr, size_t* nCols = nullptr)
      {
        py::gil_scoped_release loseGil;
        const auto toRow = toR ? *toR : (nRows ? fromR + (int)*nRows - 1 : Range::TO_END);
        const auto toCol = toC ? *toC : (nCols ? fromC + (int)*nCols - 1 : Range::TO_END);
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
      
      template<class T>
      py::object getItem(const T& range, py::object loc)
      {
        size_t fromRow, fromCol, toRow, toCol, nRows, nCols;
        std::tie(nRows, nCols) = range.shape();
        bool singleValue = getItemIndexReader2d(loc, nRows, nCols,
          fromRow, fromCol, toRow, toCol);
        return singleValue
          ? convertExcelObj(range.value((int)fromRow, (int)fromCol))
          : py::cast(range.range((int)fromRow, (int)fromCol, (int)toRow, (int)toCol));
      }

      py::object getItemSheet(const ExcelWorksheet& ws, py::object loc)
      {
        if (PyUnicode_Check(loc.ptr()))
          return py::cast((Range*)new ExcelRange(ws.range(pyToWStr(loc))));
        else
          return getItem(ws, loc);
      }

      struct RangeIter
      {
        Range& _range;
        size_t _i, _j;

        RangeIter(Range& r) : _range(r), _i(0), _j(0) 
        {}

        ExcelObj next()
        {
          if (++_j == _range.nCols())
            if (++_i == _range.nRows())
              throw py::stop_iteration();
          return _range.value(_i, _j);
        }
      };
    }


    template<class T>
    struct Collection
    {
      using value_t = decltype(T::active());
      struct Iter
      {
        vector<value_t> _objects;
        size_t i = 0;
        Iter() : _objects(T::list()) {}
        Iter(const Iter&) = delete;
        value_t next()
        {
          if (i >= _objects.size())
            throw py::stop_iteration();
          return std::move(_objects[i++]);
        }
      };
      value_t getitem(const wstring& name)
      {
        try
        {
          return value_t(name.c_str());
        }
        catch (...)
        {
          throw py::key_error();
        }
      }
      auto iter()
      {
        return new Iter();
      }
      value_t active()
      {
        return std::move(T::active());
      }
    };

    template<class T>
    auto toCom(T& p, const char* binder) 
    { 
      return comToPy(p.ptr(), binder); 
    }
    template<>
    auto toCom(Range& range, const char* binder)
    {
      return comToPy(ExcelRange(range).ptr(), binder);
    }

    static int theBinder = addBinder([](pybind11::module& mod)
    {
      py::class_<RangeIter>(mod, "RangeIter")
        .def("__iter__", [](const py::object& self) { return self; })
        .def("__next__", &RangeIter::next);

      // Bind Range class from xloil::ExcelRange
      py::class_<Range>(mod, "Range")
        .def(py::init([](const wchar_t* x) { return newRange(x); }),
          py::arg("address"))
        .def("range", subRange<Range>,
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row") = nullptr,
          py::arg("to_col") = nullptr,
          py::arg("num_rows") = nullptr,
          py::arg("num_cols") = nullptr)
        .def("cells", rangeCell,
          py::arg("row"),
          py::arg("col"))
        .def("__iter__", [](Range& self) { return new RangeIter(self); })
        .def("__getitem__", getItem<Range>)
        .def("__len__", [](const Range& r) { return r.nRows() * r.nCols(); })
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
          })
        .def("to_com", toCom<Range>, py::arg("lib") = "");

      // TODO: do we need main thread synchronisation on all this?
      py::class_<ExcelWorksheet>(mod, "Worksheet")
        .def_property_readonly("name", &ExcelWorksheet::name)
        .def_property_readonly("parent", &ExcelWorksheet::parent)
        .def("__getitem__", getItemSheet)
        .def("range", subRange<ExcelWorksheet>,
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row") = nullptr,
          py::arg("to_col") = nullptr,
          py::arg("num_rows") = nullptr,
          py::arg("num_cols") = nullptr)
        .def("at_address", (ExcelRange(ExcelWorksheet::*)(const wstring_view&) const)& ExcelWorksheet::range, py::arg("address"))
        .def("calculate", &ExcelWorksheet::calculate)
        .def("activate", &ExcelWorksheet::activate)
        .def("to_com", toCom<ExcelWorksheet>, py::arg("lib")="");

      py::class_<ExcelWorkbook>(mod, "Workbook")
        .def_property_readonly("name", &ExcelWorkbook::name)
        .def_property_readonly("path", &ExcelWorkbook::path)
        .def_property_readonly("worksheets", &ExcelWorkbook::worksheets)
        .def_property_readonly("windows", &ExcelWorkbook::windows)
        .def("worksheet", &ExcelWorkbook::worksheet, py::arg("name"))
        .def("__getitem__", &ExcelWorkbook::worksheet)
        .def("to_com", toCom<ExcelWorkbook>, py::arg("lib") = "");

      py::class_<ExcelWindow>(mod, "ExcelWindow")
        .def_property_readonly("hwnd", &ExcelWindow::hwnd)
        .def_property_readonly("name", &ExcelWindow::name)
        .def_property_readonly("workbook", &ExcelWindow::workbook)
        .def("to_com", toCom<ExcelWindow>, py::arg("lib") = "");

      using Workbooks = Collection<App::Workbooks>;
      using Windows = Collection<App::Windows>;

      py::class_<Workbooks::Iter>(mod, "ExcelWorkbooksIter")
        .def("__iter__", [](py::object self) { return self; })
        .def("__next__", &Workbooks::Iter::next);

      py::class_<Workbooks>(mod, "ExcelWorkbooks")
        .def("__getitem__", &Workbooks::getitem)
        .def("__iter__", &Workbooks::iter)
        .def_property_readonly("active", &Workbooks::active);

      py::class_<Windows::Iter>(mod, "ExcelWindowsIter")
        .def("__iter__", [](py::object self) { return self; })
        .def("__next__", &Windows::Iter::next);

      py::class_<Windows>(mod, "ExcelWindows")
        .def("__getitem__", &Windows::getitem)
        .def("__iter__", &Windows::iter)
        .def_property_readonly("active", &Windows::active);

      // Use 'new' with this return value policy or we get a segfault later. 
      mod.add_object("workbooks", py::cast(new Workbooks(), py::return_value_policy::take_ownership));
      mod.add_object("windows", py::cast(new Windows(), py::return_value_policy::take_ownership));
      mod.def("active_worksheet", &App::Worksheets::active);
      mod.def("active_workbook", &App::Workbooks::active);
    });
  }
}
