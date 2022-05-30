#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/ExcelRange.h>
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

  protected:
    IDispatch* _ptr;
    IAppObject(IDispatch* ptr = nullptr) { init(ptr); }
    void init(IDispatch* ptr);
    void assign(const IAppObject& that);
  };


  struct XLOIL_EXPORT Application : public IAppObject
  {
    Application(Excel::_Application* app);
    Application(size_t hWnd);
    //Application(const wchar_t* workbook);

    Application(Application&& that) noexcept { std::swap(_ptr, that._ptr); }
    Application(const Application& that) : IAppObject(that._ptr) {}

    Application& operator=(const Application& that) noexcept { assign(that); return *this; }
    Application& operator=(Application&& that)      noexcept { std::swap(_ptr, that._ptr); return *this; }

    Excel::_Application& com() const { return *(Excel::_Application*)_ptr; }

    virtual std::wstring name() const;

    Workbooks Workbooks() const;
    Windows Windows() const;
    ExcelWorksheet ActiveWorksheet() const;

    ExcelObj Run(const std::wstring& func, const size_t nArgs, const ExcelObj* args[]);

    void allowEvents(bool value);
  };

  /// <summary>
  /// Gets the Excel.Application object which is the root of the COM API 
  /// </summary>
  XLOIL_EXPORT Application& excelApp() noexcept;

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
    ExcelWorkbook(Excel::_Workbook* p) : IAppObject((IDispatch*)p) {}

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
    /// Returns a list of all Windows associated with this workbook (there can be
    /// multiple windows viewing a single workbook).
    /// </summary>
    /// <returns></returns>
    std::vector<ExcelWindow> windows() const;

    /// <summary>
    /// Returns a list of all Worksheets in this Workbook. It does not include 
    /// chart sheets
    /// </summary>
    /// <returns></returns>
    std::vector<ExcelWorksheet> worksheets() const;

    ExcelWorksheet worksheet(const std::wstring_view& name) const;

    /// <summary>
    /// Makes this the active workbook
    /// </summary>
    /// <returns></returns>
    void activate() const;

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
    ExcelWindow(Excel::Window* p) : IAppObject((IDispatch*)p) {}
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
    /// Gives the name of the workbook displayed by this window 
    /// </summary>
    ExcelWorkbook workbook() const;

    /// <summary>
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::Window& com() const { return *(Excel::Window*)_ptr; }
  };

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

    /// <summary>
    /// Returns a tuple (width, height)
    /// </summary>
    std::tuple<row_t, col_t> shape() const final override;

    /// <summary>
    /// Returns the tuple (top row, top column, bottom row, bottom column)  
    /// which describes the extent of the range, which is assumed rectangular.
    /// </summary>
    std::tuple<row_t, col_t, row_t, col_t> bounds() const final override;

    /// <summary>
    /// Returns the address of the range in the form
    /// 'SheetNm!A1:Z5'
    /// </summary>
    std::wstring address(bool local = false) const final override;

    /// <summary>
    /// Converts the referenced range to an ExcelObj. 
    /// References to single cells return an ExcelObj of the
    /// appropriate type. Multicell refernces return an array.
    /// </summary>
    ExcelObj value() const final override;

    /// <summary>
    /// Gets the value at (i, j) as an ExcelObj (zero-based)
    /// </summary>
    ExcelObj value(row_t i, col_t j) const final override;

    /// <summary>
    /// Sets all cells in the range to the specified value
    /// </summary>
    void set(const ExcelObj& value) final override;

    void setFormula(const std::wstring_view& formula);

    std::wstring formula() final override;

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
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
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::Range& com() const { return *(Excel::Range*)_ptr; }
    Excel::Range& com() { return *(Excel::Range*)_ptr; }
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
    ExcelWorksheet(Excel::_Worksheet* p) : IAppObject((IDispatch*)p) {}

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
    /// The raw COM ptr to the underlying object. Be sure to correctly inc ref
    /// and dec ref any use of it.
    /// </summary>
    Excel::_Worksheet& com() const { return *(Excel::_Worksheet*)_ptr; }
  };

  class XLOIL_EXPORT Workbooks
  {
  public:
    Workbooks(Application app = excelApp());
    ExcelWorkbook active() const;
    ExcelWorkbook get(const std::wstring_view& name) { return ExcelWorkbook(name, app); }
    std::vector<ExcelWorkbook> list() const;
    size_t count();

    Application app;
  };

  class XLOIL_EXPORT Windows
  {
  public:
    Windows(Application app = excelApp());
    ExcelWindow active() const;
    ExcelWindow get(const std::wstring_view& name) { return ExcelWindow(name, app); }
    std::vector<ExcelWindow> list() const;
    size_t count();

    Application app;
  };

  inline Workbooks Application::Workbooks() const
  {
    return xloil::Workbooks(*this);
  }

  inline Windows Application::Windows() const
  {
    return xloil::Windows(*this);
  }

  namespace App
  {
    struct ExcelInternals
    {
      /// <summary>
      /// The Excel major version number
      /// </summary>
      int version;
      /// <summary>
      /// The Windows API process instance handle, castable to HINSTANCE
      /// </summary>
      void* hInstance;
      /// <summary>
      /// The Windows API handle for the top level Excel window 
      /// castable to type HWND
      /// </summary>
      long long hWnd;
      size_t mainThreadId;
    };

    /// <summary>
    /// Returns Excel application state information such as the version number,
    /// HINSTANCE, window handle and thread ID.
    /// </summary>
    XLOIL_EXPORT const ExcelInternals& internals() noexcept;
  }
}