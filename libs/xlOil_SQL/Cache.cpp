#include <xloil/ObjectCache.h>
#include "Cache.h"

using std::shared_ptr;
namespace xloil 
{
  namespace SQL 
  {
    constexpr wchar_t theSqlCacheUniquifier = L'\x6B13';
    typedef ObjectCache<shared_ptr<const CacheObj>, theSqlCacheUniquifier> CacheType;
    static std::unique_ptr<CacheType> theObjCache;

    void createCache()
    {
      theObjCache.reset(new CacheType());
    }

    ExcelObj cacheAdd(shared_ptr<const CacheObj>&& obj)
    {
      if (!obj)
        return Const::Error(CellError::Value);
      return theObjCache->add(std::forward<shared_ptr<const CacheObj>>(obj));
    }
    bool cacheFetch(const std::wstring& cacheString, shared_ptr<const CacheObj>& obj)
    {
      return theObjCache->fetch(cacheString.c_str(), cacheString.length(), obj);
    }

    // TODO: another possible cache object?
    //
    //class ArrayTable : public CacheObj
    //{
    //public:
    //  virtual void addTable(std::shared_ptr<sqlite3> db, std::string name);
    //  std::shared_ptr<ExcelObj> arr;
    //  std::string schema;
    //  ;
    //};
  }
}