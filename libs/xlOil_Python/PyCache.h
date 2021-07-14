#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Caller.h>

namespace pybind11 { class object; }
namespace xloil 
{
  namespace Python 
  {
    /// <summary>
    /// Adds a python object to the cache, returning a cache reference
    /// string in an ExcelObj
    /// </summary>
    ExcelObj pyCacheAdd(const pybind11::object& obj, const wchar_t* caller = nullptr);

    /// <summary>
    /// Tries to fetch an object give a cache reference string, returning
    /// true if sucessful
    /// </summary>
    bool pyCacheGet(const std::wstring_view& cacheStr, pybind11::object& obj);

    static constexpr uint16_t CACHE_KEY_MAX_LEN = 1 + CallerInfo::INTERNAL_REF_MAX_LEN + 2;
  }
}