#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/ExcelRange.h>
#include <string>
#include <memory>
#include <vector>

struct IDispatch;
namespace Excel { struct Window; struct _Workbook; struct _Worksheet; struct Range; }

namespace xloil
{
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

  class ExcelWindow;
  class ExcelWorksheet;

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
    explicit ExcelWorkbook(const std::wstring_view& name = std::wstring_view());
    /// <summary>
    /// Constructs an ExcelWorkbook from a COM pointer
    /// </summary>
    /// <param name="p"></param>
    ExcelWorkbook(Excel::_Workbook* p) : IAppObject((IDispatch*)p) {}

    ExcelWorkbook& operator=(const ExcelWorkbook& that) { assign(that); return *this; }
    ExcelWorkbook(ExcelWorkbook&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWorkbook(const ExcelWorkbook& that) : IAppObject(that._ptr) {}
    
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


    std::vector<ExcelWorksheet> worksheets() const;

    ExcelWorksheet worksheet(const std::wstring_view& name) const;

    /// <summary>
    /// Makes this the active workbook
    /// </summary>
    /// <returns></returns>
    void activate() const;

    Excel::_Workbook* ptr() const { return (Excel::_Workbook * )_ptr; }
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
    /// <param name="windowCaption">The name of the window to find, or the active window if null</param>
    explicit ExcelWindow(const std::wstring_view& caption = std::wstring_view());

    ExcelWindow& operator=(const ExcelWindow& that) { assign(that); return *this; }
    ExcelWindow(ExcelWindow&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWindow(const ExcelWindow& that) : IAppObject(that._ptr) {}

    /// <summary>
    /// Retuns the Win32 window handle
    /// </summary>
    /// <returns></returns>
    size_t hwnd() const;
    /// <summary>
    /// Returns the window title
    /// </summary>
    /// <returns></returns>
    std::wstring name() const override;
    /// <summary>
    /// Gives the name of the workbook displayed by this window 
    /// </summary>
    ExcelWorkbook workbook() const;

    Excel::Window* ptr() const { return (Excel::Window * )_ptr; }
  };

  class XLOIL_EXPORT ExcelRange : public Range, public IAppObject
  {
  public:
    /// <summary>
    /// Constructs a Range from a address. A local address (not qualified with
    /// a workbook name) will refer to the active workbook
    /// </summary>
    explicit ExcelRange(const std::wstring_view& address);
    ExcelRange(Excel::Range* range) : IAppObject((IDispatch*)range) {}
    ExcelRange(ExcelRange&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelRange(const ExcelRange& that) : IAppObject(that._ptr) {}

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
    /// Gets the value at (i, j) as an ExcelObj
    /// </summary>
    ExcelObj value(row_t i, col_t j) const final override;

    /// <summary>
    /// Sets all cells in the range to the specified value
    /// </summary>
    void set(const ExcelObj& value) final override;

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
    void clear() final override;

    std::wstring name() const override;

    Excel::Range* ptr() const { return (Excel::Range*)_ptr; }
    Excel::Range* ptr()       { return (Excel::Range*)_ptr; }
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
    ExcelWorksheet(Excel::_Worksheet* p) : IAppObject((IDispatch*)p) {}

    //explicit ExcelWorksheet(const std::wstring_view& name = nullptr);

    ExcelWorksheet& operator=(const ExcelWorksheet& that) { assign(that); return *this; }
    ExcelWorksheet(ExcelWorksheet&& that) noexcept { std::swap(_ptr, that._ptr); }
    ExcelWorksheet(const ExcelWorksheet& that) : IAppObject(that._ptr) {}

    /// <summary>
    /// Returns the window title
    /// </summary>
    /// <returns></returns>
    std::wstring name() const override;

    /// <summary>
    /// Gives the name of the workbook which owns this sheet
    /// </summary>
    ExcelWorkbook parent() const;

    ExcelRange range(const std::wstring_view& address) const;

    ExcelRange range(int fromRow, int fromCol,
      int toRow = ExcelRange::TO_END, int toCol = ExcelRange::TO_END) const;

    ExcelRange cell(int i, int j) const
    {
      return range(i, j, i, j);
    }
    /// <summary>
    /// Convenience wrapper for cell(i,j)->value()
    /// </summary>
    ExcelObj value(Range::row_t i, Range::col_t j) const;

    std::tuple<Range::row_t, Range::col_t> shape() const { return std::make_tuple(XL_MAX_ROWS, XL_MAX_COLS); }

    void activate();
    void calculate();

    Excel::_Worksheet* ptr() const { return (Excel::_Worksheet*)_ptr; }
  };


  namespace App
  {
    struct XLOIL_EXPORT Workbooks
    {
      static ExcelWorkbook active();
      inline ExcelWorkbook get(const std::wstring_view& name) { return ExcelWorkbook(name); }
      static std::vector<ExcelWorkbook> list();
      static size_t count();
    };

    struct XLOIL_EXPORT Windows
    {
      static ExcelWindow active();
      inline ExcelWindow get(const std::wstring_view& name) { return ExcelWindow(name); }
      static std::vector<ExcelWindow> list();
      static size_t count();
    };

    struct XLOIL_EXPORT Worksheets
    {
      static ExcelWorksheet active();
    };

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