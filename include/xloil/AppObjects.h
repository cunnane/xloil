#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/Range.h>
#include <xloil/ExcelRef.h>
#include <string>
#include <memory>
#include <vector>


// Forward Declarations from Typelib
struct IDispatch;

namespace Excel 
{
  struct _Application;
  struct Window;
  struct _Workbook;
  struct _Worksheet;
  struct Range;
}

namespace xloil
{
  class ExcelWindow;
  class ExcelWorksheet;
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
  /// Base class for objects in the object model, not very usefuly directly.
  /// </summary>
  class XLOIL_EXPORT IAppObject
  {
  public:
    virtual ~IAppObject();
    /// <summary>
    /// Returns an identifier for the object. This may be a workbook name,
    /// window caption or range address.
    /// </summary>
    /// <returns></returns>
    virtual std::wstring name() const = 0;
    IDispatch* basePtr() const { return _ptr; }
    bool valid() const { return _ptr; }

    IDispatch* detach() { IDispatch* p = nullptr; std::swap(p, _ptr); return p; }

  protected:
    IDispatch* _ptr;
    IAppObject(IDispatch* ptr = nullptr, bool steal = false) { init(ptr, steal); }
    void init(IDispatch* ptr, bool steal = false);
    void assign(const IAppObject& that);
  };


  class XLOIL_EXPORT Application : public IAppObject
  {
  public:
    /// <summary>
    /// Construct an application object from a ptr to the underlying COM object.
    /// If no ptr is provided, a new instance of Excel is created.
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

    Application(Application&& that) noexcept { std::swap(_ptr, that._ptr); }
    Application(const Application& that) : IAppObject(that._ptr) {}

    Application& operator=(const Application& that) noexcept { assign(that); return *this; }
    Application& operator=(Application&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }

    Excel::_Application& com() const 
    { 
      if (!valid()) throw new NullComObjectException();
      return *(Excel::_Application*)_ptr; 
    }

    virtual std::wstring name() const;

    /// <summary>
    /// Calculates
    /// </summary>
    void calculate(const bool full=false, const bool rebuild=false);

    Workbooks Workbooks() const;
    Windows Windows() const;
    ExcelWorksheet ActiveWorksheet() const;

    ExcelObj Run(const std::wstring& func, const size_t nArgs, const ExcelObj* args[]);

    ExcelWorkbook Open(const std::wstring& filepath, bool updateLinks=true, bool readOnly=false);

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
    void setEnableEvents(bool value);
  };

  /// <summary>
  /// Gets the Excel.Application object which is the root of the COM API 
  /// </summary>
  XLOIL_EXPORT Application& excelApp();

  class XLOIL_EXPORT ExcelRange : public Range, public IAppObject
  {
  public:
    /// <summary>
    /// Constructs a Range from a address. A local address (not qualified with
    /// a workbook name) will refer to the active workbook
    /// </summary>
    explicit ExcelRange(
      const std::wstring_view& address,
      Application app = excelApp());

    ExcelRange(const Range& range);
    ExcelRange(Excel::Range* range) : IAppObject((IDispatch*)range) {}
    ExcelRange(const ExcelRef& ref) : ExcelRange(ref.address()) {}

    ExcelRange(ExcelRange&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelRange(const ExcelRange& that) : IAppObject(that._ptr) {}

    ExcelRange& operator=(ExcelRange&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }
    ExcelRange& operator=(const ExcelRange& that) noexcept { assign(that); return *this; }

    Range* range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final override;

    std::tuple<row_t, col_t> shape() const final override;

    std::tuple<row_t, col_t, row_t, col_t> bounds() const final override;

    std::wstring address(bool local = false) const final override;

    ExcelObj value() const final override;

    ExcelObj value(row_t i, col_t j) const final override;

    void set(const ExcelObj& value) final override;

    /// <summary>
    /// Sets the forumula for the range to the specified string. If the 
    /// range is larger than one cell, the formula is applied as an 
    /// ArrayFormula.
    /// </summary>
    /// <param name="formula"></param>
    void setFormula(const std::wstring_view& formula);

    /// <summary>
    /// Gets the formula assoicated with this range (or cell)
    /// </summary>
    /// <returns></returns>
    std::wstring formula() final override;

    void clear() final override;

    /// <summary>
    /// The range address
    /// </summary>
    std::wstring name() const override;

    /// <summary>
    /// The worksheet which contains this range
    /// </summary>
    ExcelWorksheet parent() const;

    /// <summary>
    /// Returns the Application object which owns this Range
    /// </summary>
    Application app() const;

    /// <summary>
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::Range& com() const { return *(Excel::Range*)_ptr; }
    Excel::Range& com() { return *(Excel::Range*)_ptr; }
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWorksheet : public IAppObject
  {
  public:
    /// <summary>
    /// Constructs an ExcelWindow from a COM pointer
    /// </summary>
    /// <param name="p"></param>
    ExcelWorksheet(Excel::_Worksheet* p, bool steal=false) : IAppObject((IDispatch*)p, steal) {}

    ExcelWorksheet(ExcelWorksheet&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWorksheet(const ExcelWorksheet& that) : IAppObject(that._ptr) {}

    ExcelWorksheet& operator=(const ExcelWorksheet& that) noexcept { assign(that); return *this; }
    ExcelWorksheet& operator=(ExcelWorksheet&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }

    /// <summary>
    /// Returns the window title
    /// </summary>
    /// <returns></returns>
    std::wstring name() const override;

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

    /// <summary>
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::_Worksheet& com() const { return *(Excel::_Worksheet*)_ptr; }
  };


  /// <summary>
  /// Wraps a Workbook (https://docs.microsoft.com/en-us/office/vba/api/excel.workbook) in
  /// Excel's object model but with very limited functionality at present
  /// </summary>
  class XLOIL_EXPORT ExcelWorkbook : public IAppObject
  {
  public:
    /// <summary>
    /// Gives the ExcelWorkbook object associated with the given workbookb name, or the active workbook
    /// </summary>
    /// <param name="name">The name of the workbook to find, or the active workbook if null</param>
    explicit ExcelWorkbook(
      const std::wstring_view& name = std::wstring_view(), 
      Application app = excelApp());
    /// <summary>
    /// Constructs an ExcelWorkbook from a COM pointer
    /// </summary>
    /// <param name="p"></param>
    ExcelWorkbook(Excel::_Workbook* p, bool steal=false) : IAppObject((IDispatch*)p, steal) {}

    ExcelWorkbook(ExcelWorkbook&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWorkbook(const ExcelWorkbook& that) : IAppObject(that._ptr) {}

    ExcelWorkbook& operator=(const ExcelWorkbook& that) noexcept { assign(that); return *this; }
    ExcelWorkbook& operator=(ExcelWorkbook&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }

    /// <inheritdoc />
    std::wstring name() const override;

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
    std::vector<ExcelWindow> windows() const;

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

    /// <summary>
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::_Workbook& com() const { return *(Excel::_Workbook*)_ptr; }
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWindow : public IAppObject
  {
  public:
    /// <summary>
    /// Constructs an ExcelWindow from a COM pointer
    /// </summary>
    /// <param name="p"></param>
    ExcelWindow(Excel::Window* p, bool steal = false) : IAppObject((IDispatch*)p, steal) {}
    /// <summary>
    /// Gives the ExcelWindow object associated with the given window caption, or the active window
    /// </summary>
    /// <param name="caption">The name of the window to find, or the active window if null</param>
    explicit ExcelWindow(
      const std::wstring_view& caption = std::wstring_view(),
      Application app = excelApp());

    ExcelWindow(ExcelWindow&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWindow(const ExcelWindow& that) : IAppObject(that._ptr) {}

    ExcelWindow& operator=(const ExcelWindow& that) noexcept { assign(that); return *this; }
    ExcelWindow& operator=(ExcelWindow&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }

    /// <summary>
    /// Retuns the Win32 window handle
    /// </summary>
    /// <returns></returns>
    size_t hwnd() const;

    /// <summary>
    /// Returns the window title
    /// </summary>
    std::wstring name() const override;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Gives the name of the workbook displayed by this window 
    /// </summary>
    ExcelWorkbook workbook() const;

    /// <summary>
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::Window& com() const { return *(Excel::Window*)_ptr; }
  };

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
    Worksheets(Application app = excelApp());
    Worksheets(ExcelWorkbook workbook);
    ExcelWorksheet active() const { return parent.app().ActiveWorksheet(); }
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

    ExcelWorkbook parent;
  };

  class XLOIL_EXPORT Workbooks
  {
  public:
    Workbooks(Application app = excelApp());
    ExcelWorkbook active() const;
    auto get(const std::wstring_view& name) const { return ExcelWorkbook(name, app); }
    auto operator[](const std::wstring_view& name) const { return get(name); };
    bool tryGet(const std::wstring_view& name, ExcelWorkbook& wb) const;
    std::vector<ExcelWorkbook> list() const;
    size_t count() const;
    ExcelWorkbook add();

    Application app;
  };

  class XLOIL_EXPORT Windows
  {
  public:
    Windows(Application app = excelApp());
    ExcelWindow active() const;
    auto get(const std::wstring_view& name) const { return ExcelWindow(name, app); }
    auto operator[](const std::wstring_view& name) const { return get(name); };
    bool tryGet(const std::wstring_view& name, ExcelWindow& window) const;
    std::vector<ExcelWindow> list() const;
    size_t count() const;

    Application app;
  };

  // Some function definitions which need to live down here due to
  // the order of declarations

  inline Workbooks Application::Workbooks() const
  {
    return xloil::Workbooks(*this);
  }

  inline Windows Application::Windows() const
  {
    return xloil::Windows(*this);
  }

  inline Worksheets ExcelWorkbook::worksheets() const
  { 
    return Worksheets(*this); 
  }
  
  inline ExcelWorksheet ExcelWorkbook::worksheet(const std::wstring_view& name) const
  {
    return worksheets().get(name);
  }
}