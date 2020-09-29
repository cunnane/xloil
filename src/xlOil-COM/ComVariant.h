#pragma once
#include <oleacc.h>
#include <xloil/ExcelObj.h>


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