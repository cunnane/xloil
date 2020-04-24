#pragma once
#include "ExcelObj.h"

namespace xloil
{
  constexpr const wchar_t theObjectCacheUnquifier = L'\x6C38';

  /// <summary>
  /// Returns true if the provided string contains the magic chars
  /// for the ExcelObj cache. Expects a counted string.
  /// </summary>
  /// <param name="str">Pointer to string start</param>
  /// <param name="length">Number of chars to read</param>
  inline bool maybeObjectCacheReference(const wchar_t* str, size_t length)
  {
    if (length < 6 || str[0] != theObjectCacheUnquifier || str[1] != L'[')
      return false;
    return true;
  }

  inline bool maybeObjectCacheReference(const ExcelObj& obj)
  {
    auto s = obj.asPascalStr();
    return maybeObjectCacheReference(s.pstr(), s.length());
  }
 
  XLOIL_EXPORT ExcelObj addCacheObject(std::shared_ptr<const ExcelObj>&& obj);

  // TODO: Could consider non const fetch in case we want to implement something like sort in-place
  // but only if we are in the same cell as object was created in
  XLOIL_EXPORT bool fetchCacheObject(const wchar_t* cacheString, size_t length, std::shared_ptr<const ExcelObj>& obj);

  inline ExcelObj addCacheObject(ExcelObj&& obj)
  {
    return addCacheObject(std::make_shared<const ExcelObj>(obj));
  }

  // TODO: registry of caches to avoid two uniquifiers
}