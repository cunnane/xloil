#pragma once
#include "PyHelpers.h"

namespace xloil
{
  class ExcelObj;
  namespace Python
  {
    PyObject* readKeywordArgs(const ExcelObj& obj);
  }
}