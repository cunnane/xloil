#pragma once
#include "CPython.h"

namespace xloil
{
  class ExcelObj;
  namespace Python
  {
    PyObject* readKeywordArgs(const ExcelObj& obj);
  }
}