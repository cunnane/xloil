#pragma once
#include "ExcelObj.h"

namespace xloil
{
  namespace detail
  {
    constexpr const wchar_t theObjectCacheUnquifier = L'\x6C38';
  }

  /// <summary>
  /// Returns true if the provided string contains the magic chars
  /// for the ExcelObj cache. Expects a counted string.
  /// </summary>
  /// <param name="str">Pointer to string start</param>
  /// <param name="length">Number of chars to read</param>
  inline bool objectCacheCheckReference(const std::wstring_view& str)
  {
    if (str.length() < 7 || str[0] != detail::theObjectCacheUnquifier || str[1] != L'[')
      return false;
    return true;
  }
  inline bool objectCacheCheckReference(const PStringView<>& pstr)
  {
    return objectCacheCheckReference(pstr.view());
  }
  inline bool objectCacheCheckReference(const ExcelObj& obj)
  {
    return objectCacheCheckReference(obj.asPascalStr());
  }
 
  /// <summary>
  /// Adds an ExcelObj to the object cache and returns an reference string
  /// (as an ExcelObj) based on the currently executing cell
  /// </summary>
  XLOIL_EXPORT ExcelObj 
    objectCacheAdd(std::shared_ptr<const ExcelObj>&& obj);

  // TODO: Could consider non const fetch in case we want to implement something like sort in-place
  // but only if we are in the same cell as object was created in
  XLOIL_EXPORT bool objectCacheFetch(
    const std::wstring_view& cacheString, std::shared_ptr<const ExcelObj>& obj);

  inline ExcelObj objectCacheAdd(ExcelObj&& obj)
  {
    return objectCacheAdd(std::make_shared<const ExcelObj>(obj));
  }

  // TODO: registry of caches to avoid two uniquifiers
}