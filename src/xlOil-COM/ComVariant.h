#pragma once
#include <xloil/ExcelObj.h>

typedef struct tagVARIANT VARIANT;

namespace xloil
{
  namespace COM
  {
    VARIANT excelObjToVariant(const ExcelObj& obj);
    ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange = false);
    VARIANT stringToVariant(const char* str);
    VARIANT stringToVariant(const wchar_t* str);
  }
}