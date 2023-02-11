#include "PyCore.h"
#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCOM.h"
#include <xlOil/AppObjects.h>

using std::shared_ptr;
using std::wstring_view;
using std::vector;
using std::wstring;
using std::string;
using std::move;
namespace py = pybind11;
using call_release_gil = py::call_guard<py::gil_scoped_release>;

namespace xloil
{
  namespace Python
  {
    PyTypeObject *theRangeType, *theXllRangeType, *theExcelRangeType;

    bool isRangeType(const PyObject* obj)
    {
      const auto* type = Py_TYPE(obj);
      return (type == theRangeType || type == theXllRangeType || type == theExcelRangeType);
    }

    namespace
    {
      auto range_Construct(const wchar_t* address)
      {
        py::gil_scoped_release noGil; 
        return new ExcelRange(address);
      }

      /// <summary>
      /// Creates a Range object in python given a helper functor. The 
      /// helper functor should not need the GIL. The reason for this function
      /// is that we do not want to hold the GIL and access Excel's COM API
      /// simutaneously 
      /// </summary>
      template<class F>
      py::object createPyRange(F&& f)
      {
        return py::cast([&] {
          py::gil_scoped_release noGil;
          return new ExcelRange(f());
        }(), py::return_value_policy::take_ownership);
      }

      // Works like the Range.Range function in VBA except is zero-based
      template <class T>
      inline auto range_subRange(const T& r,
        int fromR, int fromC,
        const py::object& toR,   const py::object& toC,
        const py::object& nRows, const py::object& nCols)
      {
        const auto toRow = !toR.is_none() ? toR.cast<int>() : (!nRows.is_none() ? fromR + nRows.cast<int>() - 1 : Range::TO_END);
        const auto toCol = !toC.is_none() ? toC.cast<int>() : (!nCols.is_none() ? fromC + nCols.cast<int>() - 1 : Range::TO_END);
        py::gil_scoped_release noGil;
        return r.range(fromR, fromC, toRow, toCol);
      }

      inline auto convertExcelObj(ExcelObj&& val)
      {
        return PySteal(PyFromAnyNoTrim()(val));
      }

      auto range_GetValue(const Range& r)
      {
        // TODO: converting Variant->ExcelObj->Python is not very efficient!
        ExcelObj val;
        {
          py::gil_scoped_release noGil;
          val = r.value();
        }
        return convertExcelObj(std::move(val));
      }

      void range_SetValue(Range& r, py::object pyVal)
      {
        // Optimise r1.value = r2
        if (isRangeType(pyVal.ptr()))
        {
          const auto& range = py::cast<const Range&>(pyVal);
          py::gil_scoped_release noGil;
          r.set(range.value());
        }
        else
        {
          // TODO: converting Python->ExcelObj->Variant is not very efficient!
          const auto val(FromPyObj()(pyVal.ptr()));
          // Must release gil when setting values in as this may trigger calcs 
          // which call back into other python functions.
          py::gil_scoped_release noGil;
          r.set(val);
        }
      };

      void range_Clear(Range& r)
      {
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

      py::object range_getItem(const Range& range, const py::object& loc)
      {
        size_t fromRow, fromCol, toRow, toCol, nRows, nCols;
        std::tie(nRows, nCols) = range.shape();
        const bool singleValue = getItemIndexReader2d(loc, nRows, nCols,
          fromRow, fromCol, toRow, toCol);
        return singleValue
          ? convertExcelObj(range.value((int)fromRow, (int)fromCol))
          : py::cast(range.range((int)fromRow, (int)fromCol, (int)toRow - 1, (int)toCol - 1));
      }

      // This is clearly not a very optimised implementation as the values
      // must be converted to / from ExcelObj and possibly Variant several times. 
      // However, it's not clear the that effort of writing ExcelObj and Variant  
      // operators for all relevant types would give worthwhile performance gains.
      auto range_InplaceArithmetic(
        const char* operation,
        py::object& self,
        const py::object& value)
      {
        auto oper = py::module::import("operator").attr(operation);
        auto& range = py::cast<Range&>(self);
        auto lhs = range_GetValue(range);
        lhs = oper(lhs, value);
        range_SetValue(range, lhs);
        return self;
      }

      // TODO: do this with templates in C++20
#define XLOIL_RANGE_OPERATOR(op) \
  [](py::object& self, const py::object& v) { return range_InplaceArithmetic(op, self, v); }

      inline Range* worksheet_subRange(const ExcelWorksheet& ws,
        int fromR, int fromC,
        const py::object& toR, const py::object& toC,
        const py::object& nRows, const py::object& nCols)
      {
        return new ExcelRange(range_subRange<ExcelWorksheet>(
          ws, fromR, fromC, toR, toC, nRows, nCols));
      }

      // We have some curious logic to avoid using the COM API whilst
      // holding the GIL.
      template<class TFinaliser>
      auto Worksheet_sliceHelper(
        const ExcelWorksheet& ws, const py::object& loc, TFinaliser finaliser)
      {
        if (PyUnicode_Check(loc.ptr()))
        {
          const auto address = to_wstring(loc);
          py::gil_scoped_release noGil;
          return finaliser(ws.range(address));
        }
        else
        {
          size_t fromRow, fromCol, toRow, toCol, nRows, nCols;
          std::tie(nRows, nCols) = ws.shape();
          getItemIndexReader2d(loc, nRows, nCols,
            fromRow, fromCol, toRow, toCol);

          py::gil_scoped_release noGil;
          return finaliser(ws.range(
            (int)fromRow, (int)fromCol, (int)toRow - 1, (int)toCol - 1));
        }
      }

      py::object worksheet_GetItem(
        const ExcelWorksheet& ws, const py::object& loc)
      {
        return Worksheet_sliceHelper(ws, loc,
          [](ExcelRange&& r) { return py::cast(r); });
      }

      void worksheet_SetItem(
        const ExcelWorksheet& ws, const py::object& loc, py::object pyValue)
      {
        ExcelObj value;

        if (isRangeType(pyValue.ptr()))
        {
          const auto& range = py::cast<const Range&>(pyValue);
          py::gil_scoped_release noGil;
          value = range.value();
        }
        else
          value = FromPyObj()(pyValue.ptr());

        Worksheet_sliceHelper(ws, loc,
          [value = move(value)](ExcelRange&& r) 
          { 
            r.set(value); 
            return 0; 
          });
      }

      py::object application_range(const Application& app, const std::wstring& address)
      {
        return createPyRange([&]() { return ExcelRange(address, app); });
      }

      py::object workbook_range(const ExcelWorkbook& wb, const std::wstring& address)
      {
        return application_range(wb.app(), address);
      }

      py::object workbook_GetItem(const ExcelWorkbook& wb, const py::object& loc)
      {
        // Somewhat unfortunately, since ExcelRange is a virtual child of the 
        // Range class declared in pybind, we need to pass a ptr to py::cast
        // which python can own, so we need to copy it (but with an rval ref)
        if (PyLong_Check(loc.ptr()))
        {
          auto i = PyLong_AsLong(loc.ptr());
          auto worksheets = wb.worksheets();
          if (i < 0 || i >= (long)worksheets.count())
            throw py::index_error();

          return py::cast(wb.worksheets().list()[i]);
        }
        else
        {
          // If loc string contains '!' it must an address. Otherwise it
          // might be a worksheet name. If it isn't that it may be a named
          // range we just pass it to Application.Range
          auto address = to_wstring(loc);
          if (address.empty())
            throw py::value_error();
          else if (address.find(L'!') != wstring::npos)
            return workbook_range(wb, address);
          else 
          {
            ExcelWorksheet ws(nullptr);
            bool isSheet;
            {
              py::gil_scoped_release noGil;
              // Remove quotes around worksheet name - these appear in
              // addresses if the sheet name contains spaces
              if (address[0] == L'\'' && address.length() > 2)
                address = address.substr(1, address.length() - 2);
              isSheet = wb.worksheets().tryGet(address, ws);
            }
            if (isSheet)
              return py::cast(ws);
            else
              return workbook_range(wb, address);
          }
        }
      }

      py::object Workbook_SetItem(
        const ExcelWorkbook& wb, const py::object& loc, py::object& value)
      {
        auto item = workbook_GetItem(wb, loc);
        auto hasSet = PyObject_GetAttrString(item.ptr(), "set");
        if (!hasSet)
          throw py::index_error("When using `workbook[X] = Y`, the index X must resolve to a Range");
        PySteal(hasSet)(value);
      }

      py::object Context_Enter(const py::object& x)
      {
        return x;
      }

      void Workbook_Exit(
        ExcelWorkbook& wb, 
        const py::object& /*exc_type*/, 
        const py::object& /*exc_val*/, 
        const py::object& /*exc_tb*/)
      {
        // Close *without* saving - saving must be done explicitly
        wb.close(false); 
      }

      void Application_Exit(
        Application& app,
        const py::object& /*exc_type*/,
        const py::object& /*exc_val*/,
        const py::object& /*exc_tb*/)
      {
        // Close *without* saving - saving must be done explicitly
        app.quit(true);
      }

      struct RangeIter
      {
        Range& _range;
        Range::row_t _i;
        Range::col_t _j;

        RangeIter(Range& r) : _range(r), _i(0), _j(0)
        {}

        auto next()
        {
          if (++_j == _range.nCols())
            if (++_i == _range.nRows())
              throw py::stop_iteration();
          return PyFromAny()(_range.value(_i, _j));
        }
      };

      template<class T>
      struct BindCollection
      {
        T _collection;

        template<class V> BindCollection(const V& x)
          : _collection(x)
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

        py::object getdefaulted(const wchar_t* name, const py::object& defaultVal)
        {
          value_t result(nullptr);
          if (_collection.tryGet(name, result))
            return py::cast(result);
          return defaultVal;
        }

        auto iter()
        {
          return new Iter(_collection);
        }

        size_t count() const { return _collection.count(); }

        bool contains(const wchar_t* name) const
        {
          value_t result(nullptr);
          return _collection.tryGet(name, result);
        }

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

        static auto startBinding(const py::module& mod, const char* name, const char* doc = nullptr)
        {
          using this_t = BindCollection<T>;

          py::class_<Iter>(mod, (string(name) + "Iter").c_str())
            .def("__iter__", [](const py::object& self) { return self; })
            .def("__next__", &Iter::next);

          return py::class_<this_t>(mod, name, doc)
            .def("__getitem__", &getitem)
            .def("__iter__", &iter)
            .def("__len__", &count)
            .def("__contains__", &contains)
            .def("get",
              &getdefaulted,
              R"(
              Tries to get the named object, returning the default if not found
            )",
              py::arg("name"),
              py::arg("default") = py::none())
            .def_property_readonly("active",
              &active,
              R"(
              Gives the active (as displayed in the GUI) object in the collection
              or None if no object has been activated.
            )");
        }
      };

      ExcelWorksheet addWorksheetToWorkbook(
        ExcelWorkbook& wb,
        const py::object& name,
        const py::object& before,
        const py::object& after)
      {
        auto cname = name.is_none() ? wstring() : to_wstring(name);
        auto cbefore = before.is_none() ? ExcelWorksheet(nullptr) : before.cast<ExcelWorksheet>();
        auto cafter = after.is_none() ? ExcelWorksheet(nullptr) : after.cast<ExcelWorksheet>();

        py::gil_scoped_release noGil;
        return wb.add(cname, cbefore, cafter);
      }

      auto addWorksheetToCollection(
        BindCollection<Worksheets>& worksheets,
        const py::object& name,
        const py::object& before,
        const py::object& after)
      {
        return addWorksheetToWorkbook(worksheets._collection.parent, name, before, after);
      }

      template<class T>
      auto toCom(T& p, const char* binder)
      {
        return comToPy(p.com(), binder);
      }
      template<>
      auto toCom(Range& range, const char* binder)
      {
        // Constructing an ExcelRange from another ExcelRange is cheap
        return comToPy(ExcelRange(range).com(), binder);
      }

      template<class T>
      auto getComAttr(T& p, const char* attrName)
      {
        return py::getattr(toCom(p, ""), attrName);
      }

      template<class T>
      void setComAttr(py::object& self, const py::object& attrName, const py::object& value)
      {
        if (PyBaseObject_Type.tp_setattro(self.ptr(), attrName.ptr(), value.ptr()) == 0)
          return;
        PyErr_Clear();
        py::setattr(toCom(py::cast<T&>(self), ""), attrName, value);
      }

      Application application_Construct(
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
          workbook = to_wstring(wbName);

        py::gil_scoped_release noGil;
        if (hWnd != 0)
          return Application(hWnd);
        else if (!workbook.empty())
          return Application(workbook.c_str());
        else
          return Application();
      }

      auto CallerInfo_Address(const CallerInfo& self, bool a1style = true)
      {
        py::gil_scoped_release noGil;
        return self.writeAddress(a1style ? AddressStyle::A1 : AddressStyle::RC);
      }

    } // namespace anon

    static int theBinder = addBinder([](py::module& mod)
    {
      static constexpr const char* toComDocString = R"(
          Returns a managed COM object which can be used to invoke Excel's full 
          object model. For details of the available calls see the Microsoft 
          documentation on the Excel Object Model. The ``lib`` used to provide COM
          support can be 'comtypes' or 'win32com'. If omitted, the default is 
          'win32com', unless specified in the XLL's ini file.
      )";
      static constexpr const char* appDocString = R"(
          Returns the parent `xloil.Application` object associated with this object.
      )";
      static constexpr const char* workbookAddDocString = R"(
        Add a worksheet, returning a `Worksheet` object.

        Parameters
        ----------
        name: str
          Names the worksheet, otherwise it will have an Excel-assigned name
        before: Worksheet
          Places the new worksheet immediately before this Worksheet object 
        after: Worksheet
          Places the new worksheet immediately before this Worksheet object.
          Specifying both `before` and `after` raises an exception.
      )";

#define XLO_CITE_API_SUFFIX(what, suffix) "See `Excel." #what " <https://docs.microsoft.com/en-us/office/vba/api/excel." #what #suffix">`_ "
#define XLO_CITE_API(what) XLO_CITE_API_SUFFIX(what, what)


      // We "forward declare" all our classes before defining their functions
      // so that when the definition happens pybind knows about the python types
      // being returned and can generate the correct docstring and type hints
      
      auto declare_Application = py::class_<Application>(mod, "Application",
        R"(
          Manages a handle to the *Excel.Application* object. This object is the root 
          of Excel's COM interface and supports a wide range of operations.

          In addition to the methods known to python, properties and methods of the 
          Application object can be resolved dynamically at runtime. The available methods
          will be familiar to VBA programmers and are well documented by Microsoft, 
          see `Object Model Overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_

          Note COM methods and properties are in UpperCamelCase, whereas python ones are 
          lower_case.

          Examples
          --------

          To get the name of the active worksheet:

          ::

              return xlo.app().ActiveWorksheet.Name

          )" XLO_CITE_API_SUFFIX(Application, (object)));

      auto declare_Range = py::class_<Range>(mod, "Range", R"(
          Represents a cell, a row, a column or a selection of cells containing a contiguous 
          blocks of cells. (Non contiguous ranges are not currently supported).
          This class allows direct access to an area on a worksheet. It uses similar 
          syntax to Excel's Range object, supporting the ``cell`` and ``range`` functions,  
          however indices are zero-based as per python's standard.

          A Range can be accessed and sliced using the usual syntax (the slice step must be 1):

          ::

              x[1, 1] # The *value* at (1, 1) as a python type: int, str, float, etc.

              x[1, :] # The second row as another Range object

              x[:-1, :-1] # A sub-range omitting the last row and column

          )" XLO_CITE_API_SUFFIX(Range, (object)));

      auto declare_Worksheet = py::class_<ExcelWorksheet>(mod, "Worksheet",
        R"(
          Allows access to ranges and properties of a worksheet. It uses similar 
          syntax to Excel's Worksheet object, supporting the ``cell`` and ``range`` functions, 
          however indices are zero-based as per python's standard.

          )" XLO_CITE_API(Worksheet));

      auto declare_Workbook = py::class_<ExcelWorkbook>(mod, "Workbook",
        R"(
          A handle to an open Excel workbook.
          )" XLO_CITE_API(Workbook));

      auto declare_Window = py::class_<ExcelWindow>(mod, "ExcelWindow",
        R"(
          A document window which displays a view of a workbook.
          )" XLO_CITE_API(Window));

      using PyWorkbooks = BindCollection<Workbooks>;
      using PyWindows = BindCollection<Windows>;
      using PyWorksheets = BindCollection<Worksheets>;

      PyWorkbooks::startBinding(mod, "Workbooks",
        R"(
          A collection of all the Workbook objects that are currently open in the 
          Excel application.  
          
          )" XLO_CITE_API(Workbooks))
        .def("add",
          [](PyWorkbooks& self) { return self._collection.add(); },
          R"(
            Creates and returns a new workbook with an Excel-assigned name
          )");

      PyWindows::startBinding(mod, "ExcelWindows",
        R"(
          A collection of all the document window objects in Excel. A document window 
          shows a view of a Workbook.

          )" XLO_CITE_API(Windows));

      PyWorksheets::startBinding(mod, "Worksheets",
        R"(
          A collection of all the Worksheet objects in the specified or active workbook. 
          
          )" XLO_CITE_API(Worksheets))
        .def("add",
          addWorksheetToCollection,
          workbookAddDocString,
          py::arg("name") = py::none(),
          py::arg("before") = py::none(),
          py::arg("after") = py::none());


      py::class_<RangeIter>(mod, "RangeIter")
        .def("__iter__", [](const py::object& self) { return self; })
        .def("__next__", &RangeIter::next);

      declare_Range
        .def(py::init(std::function(range_Construct)), 
          py::arg("address"))
        .def("range", 
          range_subRange<Range>,
          R"(
            Creates a subrange using offsets from the top left corner of the parent range.
            Like Excel's Range function, we allow negative offsets to select ranges outside the
            parent.

            Parameters
            ----------

            from_row: int
                Starting row offset from the top left of the parent range. Zero-based. Can be negative

            from_col: int
                Starting row offset from the top left of the parent range. Zero-based. Can be negative

            to_row: int
                End row offset from the top left of the parent range. This row will be *included* in 
                the range. The offset is zero-based and can be negative to select ranges outside the
                parent range. Do not specify both `to_row` and `num_rows`.

            to_col: int
                End column offset from the top left of the parent range. This column will be *included*
                in the range. The offset is zero-based and can be negative to select ranges outside 
                the parent range. Do not specify both `to_col` and `num_cols`.

            num_rows: int
                Number of rows in output range. Must be positive. If neither `num_rows` or `to_rows` 
                are specified, the range ends at the last row of the parent range.

            num_cols: int
                Number of columns in output range. Must be positive. If neither `num_cols` or `to_cols` 
                are specified, the range ends at the last column of the parent range.
          )",
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row")   = py::none(),
          py::arg("to_col")   = py::none(),
          py::arg("num_rows") = py::none(),
          py::arg("num_cols") = py::none())
        .def("cell", 
          &Range::cell,
          call_release_gil(),
          R"(
            Returns a Range object which consists of a single cell. The indices are zero-based 
            from the top left of the parent range.
          )",
          py::arg("row"),
          py::arg("col"))
        .def("trim",
          &Range::trim,
          call_release_gil(),
          R"(
            Returns a sub-range by trimming to the last non-empty (i.e. not Nil, #N/A or "") 
            row and column. The top-left remains the same so the function always returns
            at least a single cell, even if it's empty.  
          )")
        .def("__iter__", 
          [](Range& self) { return new RangeIter(self); },
          call_release_gil())
        .def("__getitem__", 
          range_getItem,
          R"(
            Given a 2-tuple, slices the range to return a sub Range or a single element.Uses
            normal python slicing conventions i.e[left included, right excluded), negative
            numbers are offset from the end.If the tuple specifies a single cell, returns
            the value in that cell, otherwise returns a Range object.
          )")
        .def("__len__", 
          [](const Range& r) { return r.nRows() * r.nCols(); },
          call_release_gil())
        .def("__str__", 
          [](const Range& r) { return r.address(false); },
          call_release_gil())
        .def("__iadd__", XLOIL_RANGE_OPERATOR("iadd"))
        .def("__isub__", XLOIL_RANGE_OPERATOR("isub"))
        .def("__imul__", XLOIL_RANGE_OPERATOR("imul"))
        .def("__itruediv__", XLOIL_RANGE_OPERATOR("itruediv"))
        .def_property("value",
          range_GetValue, range_SetValue,
          R"(
            Property which gets or sets the value for a range. A fetched value is converted
            to the most appropriate Python type using the normal generic converter.

            If you use a horizontal array for the assignment, it is duplicated down to fill 
            the entire rectangle. If you use a vertical array, it is duplicated right to fill 
            the entire rectangle. If you use a rectangular array, and it is too small for the 
            rectangular range you want to put it in, that range is padded with #N/As.
          )",
          py::return_value_policy::automatic)
        .def("set", 
          range_SetValue,
          R"(
            Sets the data in the range to the provided value. If a single value is passed
            all cells will be set to the value. If a 2d-array is provided, the array will be
            pasted at the top-left of the range with the remainging cells being set to #N/A.
            If a 1d array is provided it will be pasted at the top left and repeated down or
            right depending on orientation.
          )")
        .def("clear",
          range_Clear,
          R"(
            Clears all values and formatting.  Any cell in the range will then have Empty type.
          )")
        .def("address", 
          &Range::address,
          call_release_gil(),
          R"(
            Returns the address of the range in A1 format, e.g. *[Book]SheetNm!A1:Z5*. The 
            sheet name may be surrounded by single quote characters if it contains a space.
            If *local* is set to true, everything prior to the '!' is omitted.
          )",
          py::arg("local") = false)
        .def_property_readonly("nrows", 
          &Range::nRows,
          call_release_gil(),
          "Returns the number of rows in the range")
        .def_property_readonly("ncols", 
          &Range::nCols,
          call_release_gil(),
          "Returns the number of columns in the range")
        .def_property_readonly("shape", 
          &Range::shape,
          call_release_gil(),
          "Returns a tuple (num columns, num rows)")
        .def_property_readonly("bounds", 
          &Range::bounds,
          call_release_gil(),
          R"(
            Returns a zero-based tuple (top-left-row, top-left-col, bottom-right-row, bottom-right-col)
            which defines the Range area (currently only rectangular ranges are supported).
          )")
        .def_property("formula", 
          range_GetFormula, range_SetFormula,
          R"(
            Get / sets the forumula for the range as a string string. If the range
            is larger than one cell, the formula is applied as an ArrayFormula.
            Returns an empty string if the range does not contain a formula or array 
            formula.
          )")
        .def("to_com", 
          toCom<Range>,
          toComDocString, 
          py::arg("lib") = "")
        .def("__getattr__",
          getComAttr<Range>)
        .def("__setattr__",
          setComAttr<Range>)
        .def_property_readonly("parent", 
          [](const Range& r) { return ExcelRange(r).parent(); },
          call_release_gil(),
          "Returns the parent Worksheet for this Range");

      theRangeType = (PyTypeObject*)declare_Range.ptr();

      theExcelRangeType = (PyTypeObject*)
        py::class_<ExcelRange, Range>(mod, "_ExcelRange").ptr();

      theXllRangeType = (PyTypeObject*)
        py::class_<XllRange, Range>(mod, "_XllRange").ptr();

      declare_Worksheet
        .def("__str__", 
          &ExcelWorksheet::name,
          call_release_gil())
        .def_property_readonly("name", 
          &ExcelWorksheet::name,
          call_release_gil())
        .def_property_readonly("parent", 
          &ExcelWorksheet::parent,
          call_release_gil(),
          "Returns the parent Workbook for this Worksheet")
        .def_property_readonly("app", 
          &ExcelWorksheet::app,
          call_release_gil(),
          appDocString)
        .def("__getitem__", 
          worksheet_GetItem,
          R"(
            If the argument is a string, returns the range specified by the local address, 
            equivalent to ``at``.  
            
            If the argument is a 2-tuple, slices the sheet to return an xloil.Range.
            Uses normal python slicing conventions, i.e [left included, right excluded), negative
            numbers are offset from the end.
          )")
        .def("__setitem__", worksheet_SetItem)
        .def("range", 
          worksheet_subRange,
          R"(
            Specifies a range in this worksheet.

            Parameters
            ----------

            from_row: int
                Starting row offset from the top left of the parent range. Zero-based.

            from_col: int
                Starting row offset from the top left of the parent range. Zero-based.

            to_row: int
                End row offset from the top left of the parent range. This row will be *included* in 
                the range. The offset is zero-based. Do not specify both `to_row` and `num_rows`.

            to_col: int
                End column offset from the top left of the parent range. This column will be *included*  
                in the range. The offset is zero-based. Do not specify both `to_col` and `num_cols`.

            num_rows: int
                Number of rows in output range. Must be positive. If neither `num_rows` or `to_rows` 
                are specified, the range ends at the end of the sheet.

            num_cols: int
                Number of columns in output range. Must be positive. If neither `num_cols` or `to_cols` 
                are specified, the range ends at the end of the sheet.
          )",
          py::arg("from_row"),
          py::arg("from_col"),
          py::arg("to_row") = py::none(),
          py::arg("to_col") = py::none(),
          py::arg("num_rows") = py::none(),
          py::arg("num_cols") = py::none())
        .def("cell", 
          [](const ExcelWorksheet& self, int row, int col)
          {
            return new ExcelRange(self.cell(row, col));
          },
          call_release_gil(),
          R"(
            Returns a Range object which consists of a single cell. The indices are zero-based 
            from the top left of the parent range.
          )",
          py::arg("row"),
          py::arg("col"))
        .def("at",
          [](const ExcelWorksheet& self, const wstring& address)
          {
            return self.range(address);
          },
          call_release_gil(),
          "Returns the range specified by the local address, e.g. ``.at('B3:D6')``",
          py::arg("address"))
        .def("calculate", 
          &ExcelWorksheet::calculate,
          call_release_gil(),
          "Calculates this worksheet")
        .def("activate", 
          &ExcelWorksheet::activate,
          call_release_gil(),
          "Makes this worksheet the active sheet")
        .def("to_com", 
          toCom<ExcelWorksheet>, 
          toComDocString, 
          py::arg("lib") = "")
        .def("__getattr__",
          getComAttr<ExcelWorksheet>)
        .def("__setattr__",
          setComAttr<ExcelWorksheet>);
        
      declare_Workbook
        .def("__str__", 
          &ExcelWorkbook::name,
          call_release_gil())
        .def_property_readonly(
          "name", 
          &ExcelWorkbook::name,
          call_release_gil())
        .def_property_readonly(
          "path",
          &ExcelWorkbook::path,
          call_release_gil(),
          "The full path to the workbook, including the filename")
        .def_property_readonly("worksheets",
          [](ExcelWorkbook& wb) { return PyWorksheets(wb); },
          call_release_gil(),
          R"(
            A collection object of all worksheets which are part of this workbook
          )")
        .def_property_readonly(
          "windows",
          [](ExcelWorkbook& wb) { return PyWindows(wb); },
          call_release_gil(),
          R"(
            A collection object of all windows which are displaying this workbook
          )")
        .def_property_readonly("app", 
          &ExcelWorkbook::app,
          call_release_gil(),
          appDocString)
        .def("worksheet", 
          &ExcelWorkbook::worksheet,
          call_release_gil(),
          R"(
            Returns the named worksheet which is part of this workbook (if it exists)
            otherwise raises an exception.
          )",
          py::arg("name"))
        .def("range", 
          workbook_range, 
          "Create a `Range` object from an address such as \"Sheet!A1\" or a named range",
          py::arg("address"))
        .def("__getitem__", 
          workbook_GetItem,
          R"(
            If the index is a worksheet name, returns the `Worksheet` object,
            otherwise treats the string as a workbook address and returns a `Range`.
          )")
        .def("to_com", 
          toCom<ExcelWorkbook>, 
          toComDocString,
          py::arg("lib") = "")
        .def("__getattr__",
          getComAttr<ExcelWorkbook>)
        .def("__setattr__",
          setComAttr<ExcelWorkbook>)
        .def("add", 
          addWorksheetToWorkbook,
          workbookAddDocString,
          py::arg("name") = py::none(), 
          py::arg("before") = py::none(), 
          py::arg("after") = py::none())
        .def("save", 
          &ExcelWorkbook::save, 
          call_release_gil(),
          R"(
            Saves the Workbook, either to the specified `filepath` or if this is
            unspecified, to its original source file (an error is raised if the 
            workbook has never been saved).
          )",
          py::arg("filepath") = "")
        .def("close", 
          &ExcelWorkbook::close,
          call_release_gil(),
          R"(
            Closes the workbook. If there are changes to the workbook and the 
            workbook doesn't appear in any other open windows, the `save` argument
            specifies whether changes should be saved. If set to True, changes are 
            saved to the workbook, if False they are discared.
          )",
          py::arg("save")=true)
        .def("__enter__", Context_Enter)
        .def("__exit__", Workbook_Exit);

      declare_Window
        .def("__str__", 
          &ExcelWindow::name, 
          call_release_gil())
        .def_property_readonly("hwnd", 
          &ExcelWindow::hwnd, 
          call_release_gil(),
          "The Win32 API window handle as an integer")
        .def_property_readonly("name", 
          &ExcelWindow::name,
          call_release_gil())
        .def_property_readonly("workbook", 
          &ExcelWindow::workbook,
          call_release_gil(),
          "The workbook being displayed by this window")
        .def_property_readonly("app", 
          &ExcelWindow::app,
          call_release_gil(),
          appDocString)
        .def("to_com", 
          toCom<ExcelWindow>, 
          toComDocString, 
          py::arg("lib") = "")
        .def("__getattr__",
          getComAttr<ExcelWindow>)
        .def("__setattr__",
          setComAttr<ExcelWindow>);

      declare_Application
        .def(py::init(std::function(application_Construct)),
          R"(
            Creates a new Excel Application if no arguments are specified. Gets a handle to 
            an existing COM Application object based on the arguments.
 
            To get the parent Excel application if xlOil is embedded, used `xloil.app()`.

            Parameters
            ----------
            
            com: 
              Gets a handle to the given com object with class Excel.Appliction (marshalled 
              by `comtypes` or `win32com`).
            hwnd:
              Tries to gets a handle to the Excel.Application with given main window handle.
            workbook:
              Tries to gets a handle to the Excel.Application which has the specified workbook
              open.
          )",
          py::arg("com") = py::none(),
          py::arg("hwnd") = py::none(),
          py::arg("workbook") = py::none())
        .def_property_readonly("workbooks",
          [](Application& app) { return PyWorkbooks(app); },
          call_release_gil(),
          "A collection of all Workbooks open in this Application")
        .def_property_readonly("windows",
          [](Application& app) { return PyWindows(app); },
          call_release_gil(),
          "A collection of all Windows open in this Application")
        .def_property_readonly("workbook_paths",
          [](Application& app) { app.workbookPaths(); },
          call_release_gil(),
          "A set of the full path names of all workbooks open in this Application. "
          "Does not use COM interface.")
        .def("to_com",
          toCom<Application>,
          toComDocString,
          py::arg("lib") = "")
        .def("__getattr__", 
          getComAttr<Application>)
        .def("__setattr__",
          setComAttr<Application>)
        .def_property("visible",
          &Application::getVisible,
          [](Application& app, bool x) { app.setVisible(x); },
          call_release_gil(),
          R"(
            Determines whether the Excel window is visble on the desktop
          )")
        .def_property("enable_events",
          &Application::getEnableEvents,
          [](Application& app, bool x) { app.setEnableEvents(x); },
          call_release_gil(),
          R"(
            Pauses or resumes Excel's event handling. It can be useful when writing to a sheet
            to pause events both for performance and to prevent side effects.
          )")
        .def("range",
          application_range,
          "Create a range object from an external address such as \"[Book]Sheet!A1\"",
          py::arg("address"))
        .def("open",
          &Application::open,
          call_release_gil(),
          R"(
            Opens a workbook given its full `filepath`.

            Parameters
            ----------

            filepath: 
              path and filename of the target workbook
            update_links: 
              if True, attempts to update links to external workbooks
            read_only: 
              if True, opens the workbook in read-only mode
          )",
          py::arg("filepath"),
          py::arg("update_links") = true,
          py::arg("read_only") = false)
        .def("calculate",
          &Application::calculate,
          call_release_gil(),
          R"(
            Calculates all open workbooks

            Parameters
            ----------
            full:
              Forces a full calculation of the data in all open workbooks
            rebuild:
              For all open workbooks, forces a full calculation of the data 
              and rebuilds the dependencies. (Implies `full`)
          )",
          py::arg("full") = false,
          py::arg("rebuild") = false)
        .def("quit",
          &Application::quit,
          call_release_gil(),
          R"(
            Terminates the application. If `silent` is True, unsaved data
            in workbooks is discarded, otherwise a prompt is displayed.
          )",
          py::arg("silent") = true)
        .def("__enter__", Context_Enter)
        .def("__exit__", Application_Exit);

      py::class_<CallerInfo>(mod, 
        "Caller", R"(
          Captures the caller information for a worksheet function. On construction
          the class queries Excel via the `xlfCaller` function to determine the 
          calling cell or range. If the function was not called from a sheet (e.g. 
          via a macro), most of the methods return `None`.
        )")
        .def(py::init<>())
        .def("__str__", CallerInfo_Address)
        .def_property_readonly("sheet_name",
          [](const CallerInfo& self)
          {
            const auto name = self.sheetName();
            return name.empty() ? (py::object)py::none() : py::wstr(name);
          },
          "Gives the sheet name of the caller or None if not called from a sheet.")
        .def_property_readonly("workbook",
          [](const CallerInfo& self)
          {
            const auto name = self.workbook();
            return name.empty() ? (py::object)py::none() : py::wstr(name);
          },
          R"(
            Gives the workbook name of the caller or None if not called from a sheet.
            If the workbook has been saved, the name will contain a file extension.
          )")
        .def("address",
          CallerInfo_Address,
          R"(
            Gives the sheet address either in A1 form: '[Book]Sheet!A1' or RC form: '[Book]Sheet!R1C1'
          )",
          py::arg("a1style") = false)
        .def_property_readonly("range",
          [](const CallerInfo& self)
          {
            return createPyRange([&]() { return self.writeAddress(); });
          },
          "Range object corresponding to caller address");

      mod.def("active_worksheet", 
        []() { return excelApp().activeWorksheet(); },
        call_release_gil(),
        R"(
          Returns the currently active worksheet. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("active_workbook", 
        []() { return excelApp().workbooks().active(); },
        call_release_gil(),
        R"(
          Returns the currently active workbook. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("app", excelApp, py::return_value_policy::reference,
        R"(
          Returns the parent Excel Application object when xlOil is loaded as an
          addin. Will throw if xlOil has been imported to run automation.
        )");

      // We can only define these objects when running embedded in existing Excel
      // application. excelApp() will throw a ComConnectException if this is not
      // the case
      try
      {
        // Use 'new' with this return value policy or we get a segfault later. 
        mod.add_object("workbooks", 
          py::cast(new PyWorkbooks(excelApp()), py::return_value_policy::take_ownership));
      }
      catch (ComConnectException)
      {}
    }, 50); 
    // Up the priority of this binding to 50 as there are other places where app 
    // objects are returned and pybind needs to know the python types beforehand
  }
}
