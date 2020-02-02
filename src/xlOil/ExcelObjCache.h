#pragma once
#include "ExcelObj.h"

namespace xloil
{
  constexpr const wchar_t theObjectCacheUnquifier = L'\x6C38';

  inline bool checkObjectCacheReference(const wchar_t* str, size_t length)
  {
    if (length < 6 || str[0] != theObjectCacheUnquifier || str[1] != L'[')
      return false;
    return true;
  }

  inline bool checkObjectCacheReference(const ExcelObj& obj)
  {
    size_t len;
    auto* s = obj.asPascalStr(len);
    return checkObjectCacheReference(s, len);
  }
 
  ExcelObj addCacheObject(const std::shared_ptr<const ExcelObj>& obj);

  // TODO: Could consder non const fetch in case we want to implement something like sort in-place
  // but only if we are in the same cell as object was created in
  bool fetchCacheObject(const wchar_t* cacheString, size_t length, std::shared_ptr<const ExcelObj>& obj);

  inline ExcelObj addCacheObject(ExcelObj&& obj)
  {
    return addCacheObject(std::shared_ptr<const ExcelObj>(new ExcelObj(obj)));
  }

  // TODO: registry of caches to avoid two uniquifiers
}