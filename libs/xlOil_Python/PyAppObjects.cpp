#include "PyCore.h"
#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCOM.h"
#include "PyCore.h"
#include "PyAppCallRun.h"
#include "PyAddress.h"
#include <xlOil/AppObjects.h>
#include <xlOil/State.h>

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
      std::unordered_map<std::string, SpecialCells> theSpecialCellsEntries;

      SpecialCells specialCellsFromString(const std::string& value)
      {
        auto found = theSpecialCellsEntries.find(value);
        if (found != theSpecialCellsEntries.end())
          return found->second;
        throw py::value_error(value);
      }

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
        const auto toRow = !toR.is_none() 
          ? toR.cast<int>() 
          : (!nRows.is_none() 
            ? fromR + nRows.cast<int>() - 1 
            : Range::TO_END);
        const auto toCol = !toC.is_none() 
          ? toC.cast<int>() 
          : (!nCols.is_none() 
            ? fromC + nCols.cast<int>() - 1 
            : Range::TO_END);
        py::gil_scoped_release noGil;
        return r.range(fromR, fromC, toRow, toCol);
      }

      inline auto range_offset(const Range& r,
        int fromR, int fromC,
        const py::object& nRows, const py::object& nCols)
      {
        const auto toRow = fromR + (!nRows.is_none() ? nRows.cast<int>() - 1 : 0);
        const auto toCol = fromC + (!nCols.is_none() ? nCols.cast<int>() - 1 : 0);
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
      
      auto range_Address(Range& r, std::string& style, const bool local)
      {
        toLower(style);
        auto s = parseAddressStyle(style);
        if (local)
          s |= AddressStyle::LOCAL;
        return r.address(s);
      }

      auto range_GetFormula(Range& r)
      {
        // XllRange::formula only works from non-local functions so to 
        // minimise surpise, we convert to a COM range and call 'formula'
        ExcelObj val;
        {
          py::gil_scoped_release noGil; 
          val = ExcelRange(r).formula();
        }
        return convertExcelObj(std::move(val));
      }

      void range_SetFormula(Range& r, const py::object& pyVal)
      { 
        const auto val(FromPyObj()(pyVal.ptr()));
        py::gil_scoped_release noGil;
        ExcelRange(r).setFormula(val);
      }

      void range_SetFormulaExtra(Range& r, const py::object& pyVal, std::string& how)
      {
        const auto val(FromPyObj()(pyVal.ptr()));
        py::gil_scoped_release noGil;
        toLower(how);
        ExcelRange::SetFormulaMode mode;
        if (how == "" || how == "dynamic")
          mode = ExcelRange::DYNAMIC_ARRAY;
        else if (how == "array")
          mode = ExcelRange::ARRAY_FORMULA;
        else if (how == "implicit")
          mode = ExcelRange::OLD_ARRAY;
        else
          throw py::value_error(how);
        ExcelRange(r).setFormula(val, mode);
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

      int specialCellsValueHelper(const py::handle& values, int existing = 0)
      {
        auto p = values.ptr();
        if (PyUnicode_Check(p))
        {
          auto str = to_string(p);
          toLower(str);
          if (str.find(",") != string::npos)
          {
            XLO_THROW("Not supported");
          }
          if (str == "errors")
            return existing | int(ExcelType::Err);
          else if (str == "numbers")
            return existing | int(ExcelType::Num);
          else if (str == "logical")
            return existing | int(ExcelType::Bool);
          else if (str == "text")
            return existing | int(ExcelType::Str);
        }
        else if (PyType_Check(values.ptr()))
        {
          if ((PyTypeObject*)p == &PyUnicode_Type)
            return existing | int(ExcelType::Str);
          else if ((PyTypeObject*)p == theCellErrorType)
            return existing | int(ExcelType::Err);
          else if ((PyTypeObject*)p == &PyFloat_Type)
            return existing | int(ExcelType::Num);
          else if ((PyTypeObject*)p == &PyBool_Type)
            return existing | int(ExcelType::Bool);
        }
        else if (PyIterable_Check(values.ptr()))
        {
          auto iter = py::iter(values);
          while (iter != py::iterator::sentinel())
          {
            existing = specialCellsValueHelper(*iter, existing);
            ++iter;
          }
          return existing;
        }

        throw py::value_error("values");
      }

      py::object range_SpecialCells(
        Range& r,
        const py::object& type, 
        const py::object& values)
      {
        const auto specialCellsType = PyUnicode_Check(type.ptr())
          ? specialCellsFromString(to_string(type))
          : py::cast<SpecialCells>(type);

        int cellValues = 0;
        if (!values.is_none())
          cellValues = specialCellsValueHelper(values);

        py::gil_scoped_release noGil;
        auto result = ExcelRange(r).specialCells(
          specialCellsType, (ExcelType)cellValues);
        if (!result.valid())
          return py::none();
        else
          return py::cast(result);
      }

      py::object range_Areas(Range& r)
      {
        if (r.nAreas() == 1)
          return py::make_tuple(py::cast(r));

        // If nAreas > 1, we must already be a com range
        auto excelRange = dynamic_cast<ExcelRange*>(&r);
        if (!excelRange)
          XLO_THROW("Internal Error: unexpected range");

        vector<ExcelRange> areas;
        {
          py::gil_scoped_release noGil;
          areas = excelRange->areas().list();
        }

        return py::cast(areas);
      }

      auto range_Iter(Range& r)
      {
        auto noGil = std::make_unique<py::gil_scoped_release>();

        auto excelRange = dynamic_cast<ExcelRange*>(&r);
        if (excelRange)
        {
          auto begin = excelRange->begin();
          auto end = excelRange->end();
          noGil.reset();
          return py::make_iterator(std::move(begin), std::move(end));
        }
        auto xllRange = dynamic_cast<XllRange*>(&r);
        if (xllRange)
        {
          auto begin = xllRange->begin();
          auto end = xllRange->end();
          noGil.reset();
          return py::make_iterator(std::move(begin), std::move(end));
        }

        XLO_THROW("Range Iterator: internal error, unexpected type");
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

      auto Worksheet_sliceHelper(
        const ExcelWorksheet& ws, const py::object& loc)
      {
        if (PyUnicode_Check(loc.ptr()))
        {
          const auto address = to_wstring(loc);
          py::gil_scoped_release noGil;
          return ws.range(address);
        }
        else
        {
          size_t fromRow, fromCol, toRow, toCol, nRows, nCols;
          std::tie(nRows, nCols) = ws.shape();
          getItemIndexReader2d(loc, nRows, nCols,
            fromRow, fromCol, toRow, toCol);

          py::gil_scoped_release noGil;
          return ws.range(
            (int)fromRow, (int)fromCol, (int)toRow - 1, (int)toCol - 1);
        }
      }

      py::object worksheet_GetItem(
        const ExcelWorksheet& ws, const py::object& loc)
      {
        return py::cast(Worksheet_sliceHelper(ws, loc));
      }

      void worksheet_SetItem(
        const ExcelWorksheet& ws, 
        const py::object& loc, 
        const py::object& pyValue)
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

        auto sliced = Worksheet_sliceHelper(ws, loc);

        py::gil_scoped_release noGil;
        if (value.asStringView()._Starts_with(L"="))
          sliced.setFormula(value);
        else
          sliced.set(value);
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

      auto application_Open(
        Application& app,
        const wstring& filepath,
        bool updateLinks,
        bool readOnly,
        const py::object& delimiter)
      {
        auto delim = delimiter.is_none() ? wchar_t(0) : to_wstring(delimiter).front();
        py::gil_scoped_release noGil;
        return app.open(filepath, updateLinks, readOnly, delim);
      }

      template<class T>
      py::object castInvalidToNone(T obj)
      {
        return obj.valid() ? py::cast(obj) : py::none();
      }

      auto application_ActiveWorksheet(Application& app)
      {
        ExcelWorksheet obj;
        {
          py::gil_scoped_release noGil;
          obj = app.activeWorksheet();
        }
        return castInvalidToNone(obj);
      }

      auto application_ActiveWorkbook(Application& app)
      {
        ExcelWorkbook obj;
        {
          py::gil_scoped_release noGil;
          obj = app.workbooks().active();
        }
        return castInvalidToNone(obj);
      }

      auto application_ActiveCell(Application& app)
      {
        ExcelRange obj;
        {
          py::gil_scoped_release noGil;
          obj = app.activeCell();
        }
        return castInvalidToNone(obj);
      }

      auto application_Selection(Application& app)
      {
        ExcelRange obj;
        {
          py::gil_scoped_release noGil;
          obj = app.selection();
        }
        return castInvalidToNone(obj);
      }

      auto application_Run(Application& app, const std::wstring& func, const pybind11::args& args)
      {
        return applicationRun(app, func, args);
      }

      auto CallerInfo_Ctor()
      {
        if (!isCallerInfoSafe())
          throw py::value_error("CallerInfo is not available in this context");
        return CallerInfo();
      }

      auto CallerInfo_Address(const CallerInfo& self, std::string& style, bool local)
      {
        toLower(style);
        auto s = parseAddressStyle(style);
        if (local)
          s |= AddressStyle::LOCAL;

        py::gil_scoped_release noGil;
        return self.address(s);
      }

      template<class T>
      struct BindCollection
      {
        T _collection;

        template<class V> BindCollection(const V& x)
          : _collection(x)
        {}

        using value_t = decltype(_collection.get(wstring()));

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
          auto noGil = std::make_unique<py::gil_scoped_release>();
          auto begin = _collection.begin();
          auto end = _collection.end();
          noGil.reset();

          return py::make_iterator(std::move(begin), std::move(end));
        }

        size_t count() const { return _collection.count(); }

        bool contains(const wchar_t* name) const
        {
          value_t result(nullptr);
          return _collection.tryGet(name, result);
        }

        template<class K>
        static auto _bind(K&& klass)
        {
          using this_t = BindCollection<T>;

          return klass
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
              py::arg("default") = py::none());
        }

        static auto startBinding(
          const py::module& mod,
          const char* name,
          const char* doc = nullptr)
        {
          return _bind(
            py::class_<BindCollection<T>>(mod, name, doc));
        }
      };

      template<class T>
      struct BindCollectionWithActive : public BindCollection<T>
      {
        using BindCollection<T>::BindCollection;

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

        static auto startBinding(
          const py::module& mod,
          const char* name,
          const char* doc = nullptr)
        {
          return BindCollection<T>::_bind(
            py::class_<BindCollectionWithActive<T>>(mod, name, doc))
            .def_property_readonly("active",
              &active,
              R"(
                Gives the active (as displayed in the GUI) object in the collection
                or None if no object has been activated.
              )");
        }
      };

      ExcelWorksheet addWorksheetToWorkbook(
        const ExcelWorkbook& wb,
        const py::object& name,
        const py::object& before,
        const py::object& after)
      {
        auto cName   = name.is_none()   ? wstring() : to_wstring(name);
        auto cBefore = before.is_none() ? ExcelWorksheet(nullptr) : before.cast<ExcelWorksheet>();
        auto cAfter  = after.is_none()  ? ExcelWorksheet(nullptr) : after.cast<ExcelWorksheet>();

        py::gil_scoped_release noGil;
        return wb.add(cName, cBefore, cAfter);
      }

      auto addWorksheetToCollection(
        BindCollectionWithActive<Worksheets>& worksheets,
        const py::object& name,
        const py::object& before,
        const py::object& after)
      {
        return addWorksheetToWorkbook(
          worksheets._collection.parent(), 
          name, before, after);
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
      
      auto specialCellsEnum = py::BetterEnum<SpecialCells>(mod, "SpecialCells")
        .value("blanks", SpecialCells::Blanks, "Empty cells")
        .value("constants", SpecialCells::Constants, "Cells containing constants")
        .value("formulas", SpecialCells::Formulas, "Cells containing formulas")
        .value("last_cell", SpecialCells::LastCell, "The last cell in the used range")
        .value("comments", SpecialCells::Comments, "Cells containing notes")
        .value("visible", SpecialCells::Visible, "All visible cells")
        .value("all_format", SpecialCells::AllFormatConditions, "Cells of any format")
        .value("same_format", SpecialCells::SameFormatConditions, "Cells having the same format")
        .value("all_validation", SpecialCells::AllValidation, "Cells having validation criteria")
        .value("same_validation", SpecialCells::SameValidation, "Cells having the same validation criteria");

      theSpecialCellsEntries = specialCellsEnum.entries;

      // It used to be possible to support string conversion via 
      //    py::implicitly_convertible<std::string, Enum>()
      // But this seems to be perma-broken with the follow issue open for years:
      // https://github.com/pybind/pybind11/issues/2114
      // Leaving this note here in case a hero emerges to tackle the pybind issue.


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

          Ranges support the iterator protocol with iteration over the individual cells.
          Iteration takes place column-wise then row-wise within each range area:

          ::
              
              for cell in my_range.special_cells("constants", float):
                cell.value += 1
 
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

      using PyWorkbooks = BindCollectionWithActive<Workbooks>;
      using PyWindows = BindCollectionWithActive<Windows>;
      using PyWorksheets = BindCollectionWithActive<Worksheets>;

      PyWorkbooks::startBinding(mod, "Workbooks",
        R"(
          A collection of all the Workbook objects that are currently open in the 
          Excel application.  The collection is iterable.
          
          )" XLO_CITE_API(Workbooks))
        .def("add",
          [](PyWorkbooks& self) { return self._collection.add(); },
          R"(
            Creates and returns a new workbook with an Excel-assigned name
          )");

      PyWindows::startBinding(mod, "ExcelWindows",
        R"(
          A collection of all the document window objects in Excel. A document window 
          shows a view of a Workbook.  The collection is iterable.

          )" XLO_CITE_API(Windows));

      PyWorksheets::startBinding(mod, "Worksheets",
        R"(
          A collection of all the Worksheet objects in the specified or active workbook. 
          The collection is iterable.

          )" XLO_CITE_API(Worksheets))
        .def("add",
          addWorksheetToCollection,
          workbookAddDocString,
          py::arg("name") = py::none(),
          py::arg("before") = py::none(),
          py::arg("after") = py::none());

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
          py::arg("from_row") = 0,
          py::arg("from_col") = 0,
          py::arg("to_row")   = py::none(),
          py::arg("to_col")   = py::none(),
          py::arg("num_rows") = py::none(),
          py::arg("num_cols") = py::none())
        .def("offset",
          range_offset,
          R"(
            Similar to the *range* function, but with different defaults  

            Parameters
            ----------

            from_row: int
                Starting row offset from the top left of the parent range. Zero-based, can be negative

            from_col: int
                Starting row offset from the top left of the parent range. Zero-based, can be negative

            num_rows: int
                Number of rows in output range. Defaults to 1

            num_cols: int
                Number of columns in output range. Defaults to 1.
          )",
          py::arg("from_row"),
          py::arg("from_col"),
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
          range_Iter)
        .def("__getitem__", 
          range_getItem,
          R"(
            Given a 2-tuple, slices the range to return a sub Range or a single element.Uses
            normal python slicing conventions i.e. [left included, right excluded), negative
            numbers are offset from the end. If the tuple specifies a single cell, returns
            the value in that cell, otherwise returns a Range object.
          )")
        .def("__len__", 
          [](const Range& r) { return r.nRows() * r.nCols(); },
          call_release_gil())
        .def("__str__", 
          [](const Range& r) { return r.address(); },
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
          range_Address,
          call_release_gil(),
          R"(
            Writes the address to a string in the specified style, e.g. *[Book]SheetNm!A1:Z5*.

            Parameters
            ----------
            style: str
              The address format: "a1" or "rc". To produce an absolute / fixed addresses
              use "$a$1", "$r$c", "$a1", "a$1", etc. depending on whether you want
              both row and column to be fixed.
            local: bool
              If True, omits sheet and workbook infomation.
          )",
          py::arg("style") = "a1",
          py::arg("local") = false)
        .def("special_cells", 
          range_SpecialCells, 
          py::arg("kind"),
          py::arg("values"),
          R"(
            Returns a sub-range containg only cells of a specificed type or None if none 
            are found.  Behaves like VBA's `SpecialCells <https://learn.microsoft.com/en-us/office/vba/api/excel.range.specialcells>`_.
            The returned range is likely to be a multi-area range, you can use the `xloil.range.areas`
            property or a range iterator to step through the returned value.

            Parameters
            ----------

            kind: xloil.SpecialCells | str
              The kind of cells to return, a string value is convertered to the corresponding 
              enum value

            values: Optional[type | str | Iterable[type|str]]
              If `kind` is "constants" or "formulas", determine which types of cells to return.
              If the argument is omitted, all constants or formulas will be returned
              The argument can be one or an iterable of:
                * "errors" or *xloil.CellError*
                * "numbers" or *float*
                * "logical" or *bool*
                * "text" or *str*
          )")
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
        .def_property_readonly("row", 
          [](Range& r) { return std::get<0>(r.bounds()); },
          call_release_gil())
        .def_property_readonly("column",
          [](Range& r) { return std::get<1>(r.bounds()); },
          call_release_gil())
        .def_property_readonly("bounds", 
          &Range::bounds,
          call_release_gil(),
          R"(
            Returns a zero-based tuple (top-left-row, top-left-col, bottom-right-row, bottom-right-col)
            which defines the Range area (currently only rectangular ranges are supported).
          )")
        .def_property_readonly("areas",
          &range_Areas,
          R"(
            For a rectangular range, the property returns a 1-element list containing the range itself.
            For a multiple-area range, the property returns a list of the contigous/rectangular sub-ranges.
          )")
        .def_property("formula", 
          range_GetFormula, range_SetFormula,
          R"(
            Get / sets the formula for the range. If the cell contains a constant, this property returns 
            the value. If the cell is empty, this property returns an empty string. If the cell contains
            a formula, the property returns the formula that would be displayed in the formula bar as a
            string.  If the range is larger than one cell, the property returns an array of the values  
            which would be obtained calling `formula` on each cell.
            
            When setting, if the range is larger than one cell and a single value is passed that value
            is filled into each cell. Alternatively, you can set the formula to an array of the same 
            dimensions.
          )")
        .def("set_formula", range_SetFormulaExtra, 
          R"(
            The `how` parameter allows setting a range formula in a different way to setting its 
            `formula` property. It's unlikely you will need to use this functionality with modern
            Excel sheets. 

              * *dynamic* (or omitted): identical to setting the `formula` property
              * *array*: if the target range is larger than one cell and a single string is passed,
                set this as an array formula for the range
              * *implicit*: uses old-style implicit intersection - see "Formula vs Formula2" on MSDN

          )", 
          py::arg("formula"), 
          py::arg("how") = "")
        .def_property_readonly("has_formula", &Range::hasFormula,
          R"(
          Returns True if every cell in the range contains a formula, False if no cell
          contains a formula and None otherwise.
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
        .def("__setitem__", 
          worksheet_SetItem, 
          R"(
            Slices a range as per __getitem__. If the value being set starts with a equals sign 
            (=), the range formula is set, otherwise the value is set.  To force setting the value
            assign to `Range.value` instead.
          )")
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
          py::arg("from_row") = 0,
          py::arg("from_col") = 0,
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
        .def_property_readonly("used_range",
          &ExcelWorksheet::usedRange,
          call_release_gil(),
          "Returns a Range object that represents the used range on the worksheet")
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
          application_Open,
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
          py::arg("read_only") = false,
          py::arg("delimiter") = py::none())
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
        .def("run",
          application_Run,
          R"(
            Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
            This can call any user-defined function or macro but not built-in functions.

            The type and order of arguments expected depends on the function being called.
          )",
          py::arg("func"))
        .def_property_readonly("selection", &Application::selection)

        .def_property_readonly("active_worksheet",
          application_ActiveWorksheet,
          R"(
              Returns the currently active worksheet or None.
          )")
        .def_property_readonly("active_workbook",
          application_ActiveWorkbook,
          R"(
              Returns the currently active workbook or None.
          )")
        .def_property_readonly("active_cell",
          application_ActiveCell,
          R"(
              Returns the currently active cell as a Range or None.
          )")
        .def_property_readonly("selection",
          application_Selection,
          R"(
              Returns the currently active cell as a Range or None.
          )")
        .def_property_readonly("has_dynamic_arrays", 
          []() { return Environment::excelProcess().supportsDynamicArrays; })
        .def("__enter__", Context_Enter)
        .def("__exit__", Application_Exit);

      py::class_<CallerInfo>(mod, 
        "Caller", R"(
          Captures the caller information for a worksheet function. On construction
          the class queries Excel via the `xlfCaller` function to determine the 
          calling cell or range. If the function was not called from a sheet (e.g. 
          via a macro), most of the methods return `None`.
        )")
        .def(py::init(&CallerInfo_Ctor))
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
            Writes the address to a string in the specified style.

            Parameters
            ----------
            style: str
              The address format: "a1" or "rc". To produce an absolute / fixed addresses
              use "$a$1", "$r$c", "$a1", "a$1", etc. depending on whether you want
              both row and column to be fixed.
            local: bool
              If True, omits sheet and workbook infomation.
          )",
          py::arg("style") = "a1",
          py::arg("local") = false)
        .def_property_readonly("range",
          [](const CallerInfo& self)
          {
            return createPyRange([&]() { return self.address(); });
          },
          "Range object corresponding to caller address.  Will raise an exception if caller is not a range");

      mod.def("active_worksheet", 
        []() { return application_ActiveWorksheet(thisApp()); },
        R"(
          Returns the currently active worksheet or None. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("active_workbook", 
        []() { return application_ActiveWorkbook(thisApp()); },
        R"(
          Returns the currently active workbook or None. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("active_cell",
        []() { return application_ActiveCell(thisApp()); },
        R"(
          Returns the currently active cell as a Range or None. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("selection",
        []() { return application_Selection(thisApp()); },
        R"(
          Returns the currently selected cells as a Range or None. Will raise an exception if xlOil
          has not been loaded as an addin.
        )");

      mod.def("app", thisApp, py::return_value_policy::reference,
        R"(
          Returns the parent Excel Application object when xlOil is loaded as an
          addin. Will throw if xlOil has been imported to run automation.
        )");

      mod.def("all_workbooks", []() { return PyWorkbooks(thisApp()); },
        R"(
          Collection of workbooks for the current application. Equivalent to 
          `xloil.app().workbooks`.
        )");
    }, 50); 
    // Up the priority of this binding to 50 as there are other places where app 
    // objects are returned and pybind needs to know the python types beforehand
  }
}
