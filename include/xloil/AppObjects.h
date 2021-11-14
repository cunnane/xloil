#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/ExcelRange.h>
#include <string>
#include <memory>
#include <vector>

struct IDispatch;
namespace Excel { struct Window; struct _Workbook; struct Range; }

namespace xloil
{
  class IAppObject
  {
  public:
    virtual ~IAppObject();
    virtual std::wstring name() const = 0;
    virtual IDispatch* basePtr() const { return nullptr; }
  };

  class ExcelWindow;

  class XLOIL_EXPORT ExcelWorkbook : public IAppObject
  {
  public:
    ExcelWorkbook(Excel::_Workbook* p);
    explicit ExcelWorkbook(const wchar_t* name = nullptr);
    ExcelWorkbook(ExcelWorkbook&& that) noexcept : _wb(that._wb) { that._wb = nullptr; }
    ExcelWorkbook(const ExcelWorkbook& that) : ExcelWorkbook(that._wb) {}

    std::wstring name() const override;
    
    std::wstring path() const;
    std::vector<ExcelWindow> windows() const;

    void activate() const;

    IDispatch* basePtr() const override { return (IDispatch*)_wb; }
    Excel::_Workbook * ptr() const { return _wb; }

  private:
    Excel::_Workbook* _wb;
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWindow : public IAppObject
  {
  public:
    /// <summary>
    /// Constructs and ExcelWindow from a COM pointer
    /// </summary>
    /// <param name="p"></param>
    ExcelWindow(Excel::Window* p);
    /// <summary>
    /// Gives the ExcelWindow object associated with the give window caption, or the active window
    /// </summary>
    /// <param name="windowCaption">The name of the window to find, or the active window if null</param>
    explicit ExcelWindow(const wchar_t* windowCaption = nullptr);

    ExcelWindow(ExcelWindow&& that) noexcept : _window(that._window) { that._window = nullptr; }
    ExcelWindow(const ExcelWindow& that) : ExcelWindow(that._window) {}

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

    IDispatch* basePtr() const override { return (IDispatch*)_window; }
    Excel::Window* ptr() const { return _window; }

  private:
    Excel::Window* _window;
  };

  class XLOIL_EXPORT ExcelRange : public Range, public IAppObject
  {
  public:
    explicit ExcelRange(const wchar_t* address);
    ExcelRange(Excel::Range* range);
    ExcelRange(ExcelRange&& that) noexcept : _range(that._range) { that._range = nullptr; }
    ExcelRange(const ExcelRange& that) : ExcelRange(that._range) {}

    Range* range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final override;

    std::tuple<row_t, col_t> shape() const final override;

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

    ExcelObj value(row_t i, col_t j) const final override;

    void set(const ExcelObj& value) final override;

    /// <summary>
    /// Clears / empties all cells referred to by this ExcelRange.
    /// </summary>
    void clear() final override;

    std::wstring name() const override;

    IDispatch* basePtr() const override { return (IDispatch*)_range; }
    const Excel::Range* ptr() const { return _range; }

  private:
    Excel::Range* _range;
  };

  namespace App
  {
    XLOIL_EXPORT ExcelWorkbook activeWorkbook();
    XLOIL_EXPORT std::vector<ExcelWorkbook> workbooks();

    XLOIL_EXPORT ExcelWindow activeWindow();
    XLOIL_EXPORT std::vector<ExcelWindow> windows();

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