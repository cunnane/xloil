#pragma once
#include <xlOil/ExportMacro.h>
#include <xlOil/XlCallSlim.h>
#include <memory>
#include <string>

namespace xloil { class ExcelObj; }
namespace xloil
{
  class XLOIL_EXPORT CallerInfo
  {
  private:
    std::shared_ptr<const ExcelObj> _Address, _SheetName;
  public:
    CallerInfo();
    uint16_t addressRCLength() const;
    uint16_t writeAddress(wchar_t* buf, size_t bufLen, bool A1Style = false) const;
    std::wstring writeAddress(bool A1Style = true) const;
  };

  /// <summary>
  /// Returns the Excel A1-style column letter corresponding
  /// to a given zero-based column index
  /// </summary>
  void writeColumnName(size_t colIndex, char buf[4]);

  /// <summary>
  /// Writes a simple Excel ref including sheet name in
  /// either A1 or RxCy to  the provided string buffer. 
  /// That is, gives 'Sheet!A1' or 'Sheet!R1C1'.
  /// Returns the number of characters written
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefSheetAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style = true);

  /// <summary>
  /// Version of `xlrefSheetAddress` which returns a string rather
  /// than writing to a buffer
  /// </summary>
  XLOIL_EXPORT std::wstring xlrefSheetAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    bool A1Style = true);

  /// <summary>
  /// Similar to `xlrefSheetAddress`, but without the sheet name
  /// </summary>
  XLOIL_EXPORT std::wstring xlrefLocalAddress(
    const msxll::XLREF12& ref,
    bool A1Style = true);

  /// <summary>
  /// Writes a simple Excel ref (not including sheet name)
  /// to 'RxCy' or 'RaCy:RxCy' format in the provided string
  /// buffer. Returns the number of characters written
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefToLocalRC(
    const msxll::XLREF12& ref, 
    wchar_t* buf,
    size_t bufSize);

  /// <summary>
  /// Writes a simple Excel ref (not including sheet name)
  /// to 'A1' or 'A1:Z9' format in the provided string
  /// buffer. Returns the number of characters written.
  /// </summary>
  XLOIL_EXPORT uint16_t xlrefToLocalA1(
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize);

  /// <summary>
  /// Returns true if the user is currently in the function wizard.
  /// Quite an expensive check as Excel does not provide a built-in 
  /// way to check this.
  /// </summary>
  XLOIL_EXPORT bool inFunctionWizard();

  /// <summary>
  /// Throws "#WIZARD!" true if the user is currently in the function 
  /// wizard.  The idea being that this string will be returned to Excel
  /// by the surrounding try...catch.
  /// 
  /// Quite an expensive check as Excel does not provide a built-in 
  /// way to check this.
  /// </summary>
  inline void throwInFunctionWizard()
  {
    if (xloil::inFunctionWizard())
      throw std::runtime_error("#WIZARD!");
  }
}