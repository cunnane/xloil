#pragma once
#include <xlOil/ExcelObj.h>
#include "CPython.h"

namespace xloil
{
  class ExcelObj;
  namespace Python
  {
    PyObject* excelArrayToNestedTuple(const ExcelObj& obj);
    ExcelObj nestedIterableToExcel(const PyObject* p);
  }
}