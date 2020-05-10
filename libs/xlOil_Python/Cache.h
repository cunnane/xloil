#pragma once
#include <xlOil/ExcelObj.h>
#include "PyHelpers.h"
namespace xloil {
  namespace Python {

    void createCache();

    ExcelObj addCache(pybind11::object&& obj);

    bool fetchCache(const std::wstring_view& cacheStr, pybind11::object& obj);
  }
}