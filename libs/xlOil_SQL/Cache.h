#include <memory>
#include <string>

struct sqlite3;

namespace xloil 
{
  namespace SQL 
  {
    class CacheObj
    {
    public:
      virtual std::shared_ptr<sqlite3> getDB() const
      {
        return std::shared_ptr<sqlite3>();
      }
    };

    void 
      createCache();

    ExcelObj 
      cacheAdd(std::unique_ptr<const CacheObj>&& obj);

    bool
      cacheFetch(const std::wstring_view& cacheString, const CacheObj*& obj);
  }
}
