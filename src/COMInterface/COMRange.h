#pragma once
#include <xlOil/ExcelRange.h>
namespace Excel { struct Range; }
namespace xloil
{
  namespace COM
  {
    ExcelRange rangeFromAddress(const wchar_t* address);

    void rangeSetValue(ExcelRange& range, const ExcelObj& value);

    ExcelRange rangeFromComRange(Excel::Range* range);
  }
}