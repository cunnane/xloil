#pragma once
#include "ExcelObj.h"
#include "ExportMacro.h"

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

  void writeColumnName(size_t colIndex, char buf[4]);

  constexpr size_t CELL_ADDRESS_A1_MAX = 3 + 7 + 1 + 3 + 7 + 1;
  XLOIL_EXPORT size_t xlrefToStringA1(
    const msxll::XLREF12& ref, 
    wchar_t* buf, 
    size_t bufSize);
  XLOIL_EXPORT size_t xlrefSheetAddressA1(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style = true);
  XLOIL_EXPORT bool inFunctionWizard();
}