#pragma once
#include <xlOil/ExcelObj.h>
#include "PyHelpers.h"
namespace xloil 
{
  namespace Python 
  {
    /// <summary>
    /// Adds a python object to the cache, returning a cache reference
    /// string in an ExcelObj
    /// </summary>
    ExcelObj pyCacheAdd(pybind11::object&& obj);

    /// <summary>
    /// Tries to fetch an object give a cache reference string, returning
    /// true if sucessful
    /// </summary>
    bool pyCacheGet(const std::wstring_view& cacheStr, pybind11::object& obj);
  }
}