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
}