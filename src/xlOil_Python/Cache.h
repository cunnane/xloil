#pragma once
#include "ExcelObj.h"
#include "PyHelpers.h"
namespace xloil {
  namespace Python {

    void createCache();

    ExcelObj addCache(PyObject* obj);

    bool fetchCache(const wchar_t* cacheString, size_t length, PyObject*& obj);
  }
}