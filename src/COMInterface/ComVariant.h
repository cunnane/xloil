#pragma once
#include <oleacc.h>

namespace xloil { class ExcelObj; }

namespace xloil
{
  namespace COM
  {
    VARIANT excelObjToVariant(const ExcelObj& obj);
  }
}