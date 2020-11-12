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

    ExcelObj 
      cacheAdd(std::unique_ptr<const CacheObj>&& obj);

    const CacheObj*
      cacheFetch(const std::wstring_view& key);
  }
}
