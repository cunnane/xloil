#pragma once
#include "PyHelpers.h"

namespace xloil
{
  class ExcelObj;
  namespace Python
  {
    PyObject* excelArrayToNestedTuple(const ExcelObj& obj);
    ExcelObj nestedIterableToExcel(const PyObject* p);
  }
}