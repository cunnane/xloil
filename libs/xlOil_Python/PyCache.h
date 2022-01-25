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
    /// string in an ExcelObj. Must hold the GIL to call.
    /// </summary>
    ExcelObj pyCacheAdd(const pybind11::object& obj, const wchar_t* caller = nullptr);

    /// <summary>
    /// Tries to fetch an object give a cache reference string, returning
    /// true if sucessful. Must hold the GIL to call.
    /// </summary>
    bool pyCacheGet(const std::wstring_view& cacheStr, pybind11::object& obj);

    static constexpr uint16_t CACHE_KEY_MAX_LEN = XL_FULL_ADDRESS_RC_MAX_LEN;
  }
}