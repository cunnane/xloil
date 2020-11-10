#include <xloil/ObjectCache.h>
#include "Cache.h"

using std::unique_ptr;
namespace xloil 
{
  namespace SQL 
  {
    constexpr wchar_t theSqlCacheUniquifier = L'\x8449';
    typedef ObjectCache<unique_ptr<const CacheObj>, theSqlCacheUniquifier> CacheType;
    static std::unique_ptr<CacheType> theObjCache;

    void createCache()
    {
      theObjCache.reset(new CacheType());
    }

    ExcelObj cacheAdd(unique_ptr<const CacheObj>&& obj)
    {
      if (!obj)
        return Const::Error(CellError::Value);
      return theObjCache->add(std::forward<unique_ptr<const CacheObj>>(obj));
    }
    bool cacheFetch(const std::wstring_view& cacheString, const CacheObj*& obj)
    {
      unique_ptr<const CacheObj>* cacheObj = nullptr;
      auto ret = theObjCache->fetch(cacheString, cacheObj);
      obj = cacheObj->get();
      return ret;
    }
  }
}