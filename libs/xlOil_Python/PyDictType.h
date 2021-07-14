#pragma once
#include <Python.h>
#include <xlOil/ExcelObj.h>

namespace xloil
{
  class ExcelObj;
  namespace Python
  {
    PyObject* readKeywordArgs(const ExcelObj& obj);
  }
}