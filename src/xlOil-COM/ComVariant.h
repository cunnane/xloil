#pragma once
#include <xloil/ExcelObj.h>

typedef struct tagVARIANT VARIANT;

namespace xloil
{
  namespace COM
  {
    void excelObjToVariant(VARIANT* v, const ExcelObj& obj, bool allowRange = false);
    ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange = false);
    VARIANT stringToVariant(const char* str);
    VARIANT stringToVariant(const wchar_t* str);
    VARIANT stringToVariant(const std::wstring_view& str);
  }
}