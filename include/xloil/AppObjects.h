#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/Range.h>
#include <xloil/ExcelRef.h>
#include <string>
#include <memory>
#include <vector>
#include <set>

// Forward Declarations from Typelib
struct IDispatch;

namespace Excel 
{
  struct _Application;
  struct Window;
  struct _Workbook;
  struct _Worksheet;
  struct Range;
  struct Windows;
  struct Workbooks;
  struct Sheets;
}

namespace xloil
{
  class ExcelWindow;
  class ExcelWorksheet;
  class ExcelRange;
  class Windows;
  class Workbooks;
  class Worksheets;
  class ExcelWorkbook;
}

namespace xloil
{
  class ComConnectException : public std::runtime_error
  {
  public:
    ComConnectException(const char* message)
      : std::runtime_error(message)
    {}
  };

  /// <summary>
  /// Thrown if an AppObject has a null underlying pointer.  Currently
  /// only the Application object checks this
  /// </summary>
  class NullComObjectException : public std::exception
  {
  public:
    NullComObjectException(const char* message)
      : std::exception(message)
    {}
    NullComObjectException()
      : std::exception()
    {}
  };

  /// <summary>
  /// IDispatch ptr holder. Used internally.
  /// </summary>
  class DispatchObject
  {
  public:
    DispatchObject(IDispatch* ptr = nullptr, bool steal = false) 
    { 
      init(ptr, steal); 
    }

    DispatchObject(DispatchObject&& that) noexcept 
      : _ptr(nullptr) 
    { 
      std::swap(_ptr, that._ptr); 
    }

    DispatchObject(const DispatchObject& that) : DispatchObject(that._ptr) {}

    DispatchObject& operator=(const DispatchObject& that) noexcept 
    { 
      release();
      init(that._ptr);
      return *this; 
    }

    DispatchObject& operator=(DispatchObject&& that) noexcept 
    { 
      release(); 
      std::swap(_ptr, that._ptr); 
      return *this; 
    }

    ~DispatchObject()
    {
      release();
    }

    IDispatch* ptr() const { return _ptr; }
    bool valid() const { return _ptr; }
    void release();

  private:
    IDispatch* _ptr;
    void init(IDispatch* ptr, bool steal = false);
  };

  template <typename T, bool TCheck=false>
  class AppObject
  {
    DispatchObject _obj;

  public:
    AppObject(T* ptr = nullptr, bool steal = false) 
      : _obj((IDispatch*)ptr, steal)
    {}

    void check() const 
    {
      if constexpr (TCheck)
      {
        if (!valid()) throw new NullComObjectException();
      }
    }
    
    bool valid() const { return _obj.valid(); }
    void release() { _obj.release(); }
    auto dispatchPtr() const { return _obj.ptr(); }
    T& com() const { check(); return *(T*)_obj.ptr(); }
  };


  class XLOIL_EXPORT Application : public AppObject<Excel::_Application, true>
  {
  public:
    /// <summary>
    /// Construct an application object from a ptr to the underlying COM object.
    /// If no ptr is provided, a new instance of Excel is created. Always steals
    /// a reference, i.e. does not call AddRef.
    /// </summary>
    Application(Excel::_Application* app = nullptr);
    /// <summary>
    /// Construct an application object from a window handle to the top level
    /// Excel window (which has window class XLMAIN)
    /// </summary>
    Application(size_t hWnd);
    /// <summary>
    /// Construct an application object from a workbook name. This searches all
    /// open Excel instances for the one which has opened the workbook. It will
    /// fail if the workbook is not open
    /// </summary>
    Application(const wchar_t* workbook);

    std::wstring name() const;

    /// <summary>
    /// Calculates
    /// </summary>
    void calculate(const bool full=false, const bool rebuild=false);

    Workbooks workbooks() const;
    Windows windows() const;
    ExcelWorksheet activeWorksheet() const;

    ExcelObj run(const std::wstring& func, const size_t nArgs, const ExcelObj* args[]);

    ExcelWorkbook open(
      const std::wstring& filepath, 
      bool updateLinks=true, 
      bool readOnly=false,
      wchar_t delimiter = 0);

    /// <summary>
    /// The set of full path names of all open workbooks
    /// </summary>
    const std::set<std::wstring>& workbookPaths();

    /// <summary>
    /// Calls Application.Quit to close the Excel instance and frees the COM resources.
    /// This invalidates the Application object: any further calls to methods other 
    /// than quit() will raise an exception.
    /// </summary>
    /// <param name="silent">
    ///   If true, supresses save file dialogs: unsaved changes to workbooks will be discarded.
    /// </param>
    void quit(bool silent=true);

    bool getVisible() const;
    void setVisible(bool x);

    bool getEnableEvents();
    bool setEnableEvents(bool value);

    bool getDisplayAlerts();
    bool setDisplayAlerts(bool value);

    /// <summary>
    /// Returns an invalid ExcelRange is the selection is not a range
    /// </summary>
    ExcelRange selection();
  };

  /// <summary>
  /// Gets the Excel.Application object which is the root of the COM API 
  /// </summary>
  XLOIL_EXPORT Application& excelApp();

  class XLOIL_EXPORT ExcelRange : public Range, public AppObject<Excel::Range>
  {
  public:
    /// <summary>
    /// Constructs a Range from a address. A local address (not qualified with
    /// a workbook name) will refer to the active workbook
    /// </summary>
    explicit ExcelRange(
      const std::wstring_view& address,
      const Application& app = excelApp());

    ExcelRange(const Range& range);
    ExcelRange(const ExcelRef& ref) : ExcelRange(ref.address()) {}

    using AppObject<Excel::Range>::AppObject;

    std::unique_ptr<Range> range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final override;

    std::unique_ptr<Range> trim() const final override;

    std::tuple<row_t, col_t> shape() const final override;

    std::tuple<row_t, col_t, row_t, col_t> bounds() const final override;

    std::wstring address(bool local = false) const final override;

    ExcelObj value() const final override;

    ExcelObj value(row_t i, col_t j) const final override;

    void set(const ExcelObj& value) final override;

    std::wstring formula() const final override;

    void clear() final override;

    virtual Excel::Range* asComPtr() const final override
    {
      return &com();
    }

    /// <summary>
    /// Sets the forumula for the range to the specified string. If the 
    /// range is larger than one cell, the formula is applied as an 
    /// ArrayFormula.
    /// </summary>
    /// <param name="formula"></param>
    void setFormula(const std::wstring_view& formula);

    /// <summary>
    /// The range address
    /// </summary>
    std::wstring name() const;

    /// <summary>
    /// The worksheet which contains this range
    /// </summary>
    ExcelWorksheet parent() const;

    /// <summary>
    /// Returns the Application object which owns this Range
    /// </summary>
    Application app() const;
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWorksheet : public AppObject<Excel::_Worksheet>
  {
  public:
    using AppObject<Excel::_Worksheet>::AppObject;

    /// <summary>
    /// Returns the window title
    /// </summary>
    /// <returns></returns>
    std::wstring name() const;

    /// <summary>
    /// Gives the name of the workbook which owns this sheet
    /// </summary>
    ExcelWorkbook parent() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Returns a range on this worksheet given a (local) address
    /// </summary>
    ExcelRange range(const std::wstring_view& address) const;

    /// <summary>
    /// Returns the range on this worksheet starting and ending at the specified
    /// rows and columns.  Includes the start row and columns and the end row and 
    /// column, just as Excel's Worksheet.Range behaves.
    /// </summary>
    ExcelRange range(int fromRow, int fromCol,
      int toRow = ExcelRange::TO_END, int toCol = ExcelRange::TO_END) const;

    /// <summary>
    /// Returns a Range referring to the single cell (i, j) (zero-based indexing)
    /// </summary>
    ExcelRange cell(int i, int j) const
    {
      return range(i, j, i, j);
    }

    /// <summary>
    /// Convenience wrapper for cell(i,j)->value()
    /// </summary>
    ExcelObj value(Range::row_t i, Range::col_t j) const;

    /// <summary>
    /// Returns a range object representing the range used in this worksheet. It 
    /// is bounded by the top-left and the bottom right-used cells.
    /// </summary>
    ExcelRange usedRange() const;
    /// <summary>
    /// Returns the size of the worksheet, which is always (MaxRows, MaxCols).
    /// This function exists mainly to provide some polymorphism with Range.
    /// </summary>
    std::tuple<Range::row_t, Range::col_t> shape() const
    {
      return std::make_tuple(XL_MAX_ROWS, XL_MAX_COLS);
    }

    /// <summary>
    /// Makes this worksheet the active one in its workbook
    /// </summary>
    void activate();

    /// <summary>
    /// Calculates this worksheet
    /// </summary>
    void calculate();

    /// <summary>
    /// Sets the worksheet name (note 31 char limit)
    /// </summary>
    void setName(const std::wstring_view& name);
  };


  /// <summary>
  /// Wraps a Workbook (https://docs.microsoft.com/en-us/office/vba/api/excel.workbook) in
  /// Excel's object model but with very limited functionality at present
  /// </summary>
  class XLOIL_EXPORT ExcelWorkbook : public AppObject<Excel::_Workbook>
  {
  public:
    /// <summary>
    /// Gives the ExcelWorkbook object associated with the given workbookb name, or the active workbook
    /// </summary>
    /// <param name="name">The name of the workbook to find, or the active workbook if null</param>
    explicit ExcelWorkbook(
      const std::wstring_view& name = std::wstring_view(), 
      Application app = excelApp());
    
    using AppObject<Excel::_Workbook>::AppObject;

    std::wstring name() const;

    /// <summary>
    /// Returns the full file path and file name for this workbook
    /// </summary>
    std::wstring path() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Returns a list of all Windows associated with this workbook (there can be
    /// multiple windows viewing a single workbook).
    /// </summary>
    /// <returns></returns>
    Windows windows() const;

    /// <summary>
    /// Returns a collection of all Worksheets in this Workbook. It does not  
    /// include chart sheets (or Excel4 macro sheets)
    /// </summary>
    /// <returns></returns>
    Worksheets worksheets() const;

    /// <summary>
    /// Returns a worksheet object corresponding to the specified name if 
    /// it exists
    /// </summary>
    ExcelWorksheet worksheet(const std::wstring_view& name) const;

    /// <summary>
    /// Returns a range in this workbook given an address
    /// </summary>
    ExcelRange range(const std::wstring_view& address) const
    {
      return ExcelRange(address, app());
    }

    /// <summary>
    /// Adds a new worksheet, naming it if a name is provided, otherwise it
    /// will have a default name provided by Excel, such as 'Sheet4'.
    /// </summary>
    /// <param name="name"></param>
    /// <param name="before"></param>
    /// <param name="after"></param>
    /// <returns></returns>
    ExcelWorksheet add(
        const std::wstring_view& name = std::wstring_view(),
        const ExcelWorksheet& before = ExcelWorksheet(nullptr),
        const ExcelWorksheet& after = ExcelWorksheet(nullptr)) const;

    /// <summary>
    /// Makes this the active workbook
    /// </summary>
    /// <returns></returns>
    void activate() const;

    void save(const std::wstring_view& filepath = std::wstring_view());

    void close(bool save=true);
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWindow : public AppObject<Excel::Window>
  {
  public:
    using AppObject<Excel::Window>::AppObject;

    /// <summary>
    /// Gives the ExcelWindow object associated with the given window caption, or the active window
    /// </summary>
    /// <param name="caption">The name of the window to find, or the active window if null</param>
    explicit ExcelWindow(
      const std::wstring_view& caption = std::wstring_view(),
      Application app = excelApp());

    /// <summary>
    /// Retuns the Win32 window handle
    /// </summary>
    /// <returns></returns>
    size_t hwnd() const;

    /// <summary>
    /// Returns the window title
    /// </summary>
    std::wstring name() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Gives the name of the workbook displayed by this window 
    /// </summary>
    ExcelWorkbook workbook() const;
  };

  inline std::wstring to_wstring(const ExcelRange& x) { return x.name(); }

  XLOIL_EXPORT ExcelRef refFromComRange(Excel::Range& range);

  inline ExcelRef refFromRange(const Range& range)
  {
    auto excelRange = dynamic_cast<const ExcelRange*>(&range);
    if (excelRange)
      return refFromComRange(excelRange->com());
    else
      return static_cast<const XllRange&>(range).asRef();
  }

  class XLOIL_EXPORT Worksheets
  {
  public:
    Worksheets(const Application& app = excelApp());
    Worksheets(const ExcelWorkbook& workbook);
    ExcelWorksheet active() const { return app().activeWorksheet(); }
    ExcelWorksheet get(const std::wstring_view& name) const;
    auto operator[](const std::wstring_view& name) const { return get(name); };
    bool tryGet(const std::wstring_view& name, ExcelWorksheet& wb) const;
    std::vector<ExcelWorksheet> list() const;
    size_t count() const;
    ExcelWorksheet add(
      const std::wstring_view& name = std::wstring_view(),
      const ExcelWorksheet& before = ExcelWorksheet(nullptr),
      const ExcelWorksheet& after = ExcelWorksheet(nullptr)) const
    {
      return parent.add(name, before, after);
    }
    Application app() const { return parent.app(); }
    ExcelWorkbook parent;
  };

  class XLOIL_EXPORT Workbooks : public AppObject<Excel::Workbooks>
  {
  public:
    Workbooks(const Application& app = excelApp());
    ExcelWorkbook active() const;
    auto get(const std::wstring_view& name) const { return ExcelWorkbook(name, app()); }
    auto operator[](const std::wstring_view& name) const { return get(name); };
    bool tryGet(const std::wstring_view& name, ExcelWorkbook& wb) const;
    std::vector<ExcelWorkbook> list() const;
    size_t count() const;
    ExcelWorkbook add();

    Application app() const;
  };

  class XLOIL_EXPORT Windows : public AppObject<Excel::Windows>
  {
  public:
    Windows(const Application& app = excelApp());
    Windows(const ExcelWorkbook& workbook);
    ExcelWindow active() const;
    auto get(const std::wstring_view& name) const { return ExcelWindow(name, app()); }
    auto operator[](const std::wstring_view& name) const { return get(name); };
    bool tryGet(const std::wstring_view& name, ExcelWindow& window) const;
    std::vector<ExcelWindow> list() const;
    size_t count() const;

    Application app() const;
  };


  // Some function definitions which need to live down here due to
  // the order of declarations

  inline Workbooks Application::workbooks() const
  {
    return Workbooks(*this);
  }

  inline Windows Application::windows() const
  {
    return Windows(*this);
  }

  inline Worksheets ExcelWorkbook::worksheets() const
  { 
    return Worksheets(*this); 
  }

  inline Windows ExcelWorkbook::windows() const
  {
    return Windows(*this);
  }

  inline ExcelWorksheet ExcelWorkbook::worksheet(const std::wstring_view& name) const
  {
    return worksheets().get(name);
  }
}