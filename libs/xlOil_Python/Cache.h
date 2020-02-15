#pragma once
#include "ExcelObj.h"
#include "PyHelpers.h"
namespace xloil {
  namespace Python {

    void createCache();

    ExcelObj addCache(pybind11::object&& obj);

    bool fetchCache(const wchar_t* cacheString, size_t length, pybind11::object& obj);
  }
}