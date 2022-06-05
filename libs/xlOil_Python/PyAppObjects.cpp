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
using std::string;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    PyTypeObject* rangeType;

    namespace
    {
      Range* range_Construct(const wchar_t* address) 
      {
        py::gil_scoped_release noGil; 
        return new ExcelRange(address);
      }

      // Works like the Range.Range function in VBA except is zero-based
      template <class T>
      inline auto range_subRange(const T& r,
        int fromR, int fromC,
        const py::object& toR, const py::object& toC,
        const py::object& nRows, const py::object& nCols)
      {
        py::gil_scoped_release noGil;
        const auto toRow = !toR.is_none() ? toR.cast<int>() : (!nRows.is_none() ? fromR + nRows.cast<int>() - 1 : Range::TO_END);
        const auto toCol = !toC.is_none() ? toC.cast<int>() : (!nCols.is_none() ? fromC + nCols.cast<int>() - 1 : Range::TO_END);
        return r.range(fromR, fromC, toRow, toCol);
      }

      inline Range* worksheet_subRange(const ExcelWorksheet& ws,
        int fromR, int fromC,
        const py::object& toR, const py::object& toC,
        const py::object& nRows, const py::object& nCols)
      {
        return new ExcelRange(range_subRange<ExcelWorksheet>(ws, fromR, fromC, toR, toC, nRows, nCols));
      }

      inline auto convertExcelObj(ExcelObj&& val)
      {
        return PySteal(PyFromAny()(val));
      }

      auto range_GetValue(const Range& r)
      {
        return convertExcelObj(r.value());
      }

      void range_SetValue(Range& r, const py::object& pyval)
      {
        const auto val = FromPyObj()(pyval.ptr());
        // Release gil when setting values in as this may trigger calcs 
        // which call back into other python functions.
        py::gil_scoped_release noGil;
        r = val;
      };

      void range_Clear(Range& r)
      {
        // Release gil - see reasons above
        py::gil_scoped_release noGil;
        r.clear();
      }
      
      auto range_GetFormula(Range& r)
      {
        // XllRange::formula only works from non-local functions so to 
        // minimise surpise, we convert to a COM range and call 'formula'
        py::gil_scoped_release noGil;
        return ExcelRange(r).formula();
      }

      void range_SetFormula(Range& r, const wstring& val) 
      { 
        py::gil_scoped_release noGil;
        ExcelRange(r).setFormula(val);
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
          : py::cast(range.range((int)fromRow, (int)fromCol, (int)toRow - 1, (int)toCol - 1));
      }

      py::object worksheet_GetItem(const ExcelWorksheet& ws, py::object loc)
      {
        // Somewhat unfortunately, since ExcelRange is a virtual child of the 
        // Range class declared in pybind, we need to pass a ptr to py::cast
        // which python can own, so we need to copy it (but with an rval ref)
        if (PyUnicode_Check(loc.ptr()))
          return py::cast((Range*)new ExcelRange(ws.range(pyToWStr(loc))));
        else
          return getItem(ws, loc);
      }

      struct RangeIter
      {
        Range& _range;
        Range::row_t _i;
        Range::col_t _j;

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
      T _collection;

      Collection(Application app)
        : _collection(app)
      {}

      using value_t = decltype(_collection.active());
      struct Iter
      {
        vector<value_t> _objects;
        size_t i = 0;
        Iter(const T& collection) : _objects(collection.list()) {}
        Iter(const Iter&) = delete;
        value_t next()
        {
          if (i >= _objects.size())
            throw py::stop_iteration();
          return std::move(_objects[i++]);
        }
      };
      
      auto getitem(const wstring& name)
      {
        try
        {
          py::gil_scoped_release noGil;
          return _collection.get(name.c_str());
        }
        catch (...)
        {
          throw py::key_error();
        }
      }
      
      py::object getdefaulted(const wstring& name, const py::object& defaultVal)
      {
        value_t result(nullptr);
        if (_collection.tryGet(name.c_str(), result))
          return py::cast(result);
        return defaultVal;
      }

      auto iter()
      {
        return new Iter(_collection);
      }

      size_t count() const { return _collection.count(); }

      py::object active()
      {
        value_t obj(nullptr);
        {
          py::gil_scoped_release noGil;
          obj = _collection.active();
        }
        if (!obj.valid())
          return py::none();
        return py::cast(std::move(obj));
      }

      static auto startBinding(const py::module& mod, const char* name)
      {
        using this_t = Collection<T>;

        py::class_<Iter>(mod, (string(name) + "Iter").c_str())
          .def("__iter__", [](const py::object& self) { return self; })
          .def("__next__", &Iter::next);

        return py::class_<this_t>(mod, name)
          .def("__getitem__", &getitem)
          .def("__iter__", &iter)
          .def("__len__", &count)
          .def("get", &getdefaulted, py::arg("name"), py::arg("default") = py::none())
          .def_property_readonly("active", &active);
      }
    };

    template<class T>
    auto toCom(T& p, const char* binder) 
    { 
      return comToPy(p.com(), binder); 
    }
    template<>
    auto toCom(Range& range, const char* binder)
    {
      return comToPy(ExcelRange(range).com(), binder);
    }

    Application createExcelApp(
      const py::object& com,
      const py::object& hwnd,
      const py::object& wbName)
    {
      size_t hWnd = 0;
      wstring workbook;

      if (!com.is_none())
      {
        // TODO: we could get the underlying COM ptr depending on use of comtypes/pywin32
        hWnd = py::cast<size_t>(com.attr("hWnd")());
      }
      else if (!hwnd.is_none())
        hWnd = py::cast<size_t>(hwnd);
      else if (!wbName.is_none())
        workbook = pyToWStr(wbName);

      py::gil_scoped_release noGil;
      if (hWnd != 0)
        return Application(hWnd);
      else if (!workbook.empty())
        return Application(workbook.c_str());
      else
        return Application();
    }

    static int theBinder = addBinder([](pybind11::module& mod)
    {
      py::class_<RangeIter>(mod, "RangeIter")
        .def("__iter__", [](const py::object& self) { return self; })
        .def("__next__", &RangeIter::next);

      // Bind Range class from xloil::ExcelRange
      auto rangeClass = py::class_<Range>(mod, "Range")
        .def(py::init(std::function(range_Construct)), 
          py::arg("address"))
        .def("range", range_subRange<Range>,
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row")   = py::none(),
          py::arg("to_col")   = py::none(),
          py::arg("num_rows") = py::none(),
          py::arg("num_cols") = py::none())
        .def("cell", wrapNoGil(&Range::cell),
          py::arg("row"),
          py::arg("col"))
        .def("__iter__", [](Range& self) { return new RangeIter(self); })
        .def("__getitem__", getItem<Range>)
        .def("__len__", [](const Range& r) { return r.nRows() * r.nCols(); })
        .def_property("value",
          range_GetValue,
          range_SetValue,
          py::return_value_policy::automatic)
        .def("set", range_SetValue)
        .def("clear", range_Clear)
        .def("address", wrapNoGil(&Range::address), py::arg("local") = false)
        .def_property_readonly("nrows", &Range::nRows)
        .def_property_readonly("ncols", &Range::nCols)
        .def_property_readonly("shape", &Range::shape)
        .def_property("formula", range_GetFormula, range_SetFormula)
        .def("to_com", toCom<Range>, py::arg("lib") = "");

      rangeType = (PyTypeObject*)rangeClass.ptr();

      // TODO: do we need main thread synchronisation on all this?
      py::class_<ExcelWorksheet>(mod, "Worksheet")
        .def_property_readonly("name", wrapNoGil(&ExcelWorksheet::name))
        .def_property_readonly("parent", wrapNoGil(&ExcelWorksheet::parent))
        .def("__getitem__", worksheet_GetItem)
        .def("range", worksheet_subRange,
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row") = py::none(),
          py::arg("to_col") = py::none(),
          py::arg("num_rows") = py::none(),
          py::arg("num_cols") = py::none())
        .def("cell", wrapNoGil(&ExcelWorksheet::cell),
          py::arg("row"),
          py::arg("col"))
        .def("at", 
          wrapNoGil((ExcelRange(ExcelWorksheet::*)(const wstring_view&) const) &ExcelWorksheet::range),
          py::arg("address"))
        .def("calculate", wrapNoGil(&ExcelWorksheet::calculate))
        .def("activate", wrapNoGil(&ExcelWorksheet::activate))
        .def("to_com", toCom<ExcelWorksheet>, py::arg("lib")="");

      py::class_<ExcelWorkbook>(mod, "Workbook")
        .def_property_readonly("name", wrapNoGil(&ExcelWorkbook::name))
        .def_property_readonly("path", wrapNoGil(&ExcelWorkbook::path))
        .def_property_readonly("worksheets", wrapNoGil(&ExcelWorkbook::worksheets))
        .def_property_readonly("windows", wrapNoGil(&ExcelWorkbook::windows))
        .def("worksheet", wrapNoGil(&ExcelWorkbook::worksheet), py::arg("name"))
        .def("__getitem__", wrapNoGil(&ExcelWorkbook::worksheet))
        .def("to_com", toCom<ExcelWorkbook>, py::arg("lib") = "");

      py::class_<ExcelWindow>(mod, "ExcelWindow")
        .def_property_readonly("hwnd", wrapNoGil(&ExcelWindow::hwnd))
        .def_property_readonly("name", wrapNoGil(&ExcelWindow::name))
        .def_property_readonly("workbook", wrapNoGil(&ExcelWindow::workbook))
        .def("to_com", toCom<ExcelWindow>, py::arg("lib") = "");

      using PyWorkbooks  = Collection<Workbooks>;
      using PyWindows    = Collection<Windows>;
      using PyWorksheets = Collection<Worksheets>;

      py::class_<Application>(mod, "Application")
        .def(py::init(std::function(createExcelApp)),
          py::arg("com") = py::none(),
          py::arg("hwnd") = py::none(),
          py::arg("workbook") = py::none())
        .def("workbooks", [](Application& app) { py::gil_scoped_release noGil; return PyWorkbooks(app); })
        .def("windows", [](Application& app) { py::gil_scoped_release noGil; return PyWindows(app); })
        .def("to_com", toCom<Application>, py::arg("lib") = "")
        .def_property("visible",
          [](Application& app) { py::gil_scoped_release noGil; return app.getVisible(); },
          [](Application& app, bool x) { py::gil_scoped_release noGil; app.setVisible(x); })
        .def_property("enable_events",
          [](Application& app) { py::gil_scoped_release noGil; return app.getEnableEvents(); },
          [](Application& app, bool x) { py::gil_scoped_release noGil; app.setEnableEvents(x); })
        .def("quit", wrapNoGil(&Application::Quit));


      PyWorkbooks::startBinding(mod, "Workbooks")
        .def("add", [](PyWorkbooks& self) { self._collection.add(); });

      PyWindows::startBinding(mod, "ExcelWindows");

      PyWorksheets::startBinding(mod, "Worksheets");

      try
      {
        // Use 'new' with this return value policy or we get a segfault later. 
        mod.add_object("workbooks", py::cast(new PyWorkbooks(excelApp()), py::return_value_policy::take_ownership));
        mod.add_object("windows", py::cast(new PyWindows(excelApp()), py::return_value_policy::take_ownership));
        mod.def("active_worksheet", []() { return excelApp().ActiveWorksheet(); });
        mod.def("active_workbook", []() { return excelApp().Workbooks().active(); });
      }
      catch (ComConnectException)
      {
      }
    });
  }
}
