#pragma once
#include <xlOil/ExportMacro.h>
#include <xlOil/XlCallSlim.h>
#include <memory>

namespace xloil { class ExcelObj; }
namespace xloil
{
  class XLOIL_EXPORT CallerInfo
  {
  private:
    std::shared_ptr<const ExcelObj> _Address, _SheetName;
  public:
    CallerInfo();
    size_t fullAddressLength() const;
    size_t writeFullAddress(wchar_t* buf, size_t bufLen) const;
  };

  /// <summary>
  /// Returns the Excel A1-style column letter corresponding
  /// to a given zero-based column index
  /// </summary>
  void writeColumnName(size_t colIndex, char buf[4]);

  /// <summary>
  /// Writes a simple Excel ref (not including sheet name)
  /// to 'A1' or 'A1:Z9' format in the provided string
  /// buffer. Returns the number of characters written
  /// </summary>
  XLOIL_EXPORT size_t xlrefToStringA1(
    const msxll::XLREF12& ref, 
    wchar_t* buf, 
    size_t bufSize);

  /// <summary>
  /// Writes a simple Excel ref including sheet name in
  /// either A1 or RxCy to  the provided string buffer. 
  /// That is, gives 'Sheet!A1' or 'Sheet!R1C1'.
  /// Returns the number of characters written
  /// </summary>
  XLOIL_EXPORT size_t xlrefSheetAddressA1(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style = true);

  /// <summary>
  /// Writes a simple Excel ref (not including sheet name)
  /// to 'RxCy' or 'RaCy:RxCy' format in the provided string
  /// buffer. Returns the number of characters written
  /// </summary>
  XLOIL_EXPORT size_t xlrefToStringRC(
    const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize);

  bool inFunctionWizard();
}