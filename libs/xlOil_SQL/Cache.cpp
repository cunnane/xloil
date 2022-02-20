#include <xloil/ExcelObjCache.h>
#include "Cache.h"

using std::unique_ptr;
namespace xloil 
{
  template<>
  struct CacheUniquifier<std::unique_ptr<const SQL::CacheObj>>
  {
    static constexpr wchar_t value = L'\x8449';
    static constexpr auto tag = L"SQLDB";
  };

  namespace SQL
  {
    ExcelObj cacheAdd(unique_ptr<const CacheObj>&& obj)
    {
      if (!obj)
        return Const::Error(CellError::Value);
      return ExcelObj(addCached<CacheObj>(obj.release()));
    }
    const CacheObj* cacheFetch(const std::wstring_view& key)
    {
      return getCached<CacheObj>(key);
    }
  }
}