
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include <xloil/Caller.h>
#include <xlOil/ExcelArray.h>
#include "Common.h"
#include "Cache.h"

using std::shared_ptr;
using std::vector;
using std::make_unique;

namespace xloil
{
  namespace SQL
  {
    class DataBaseRef : public CacheObj
    {
    public:
      DataBaseRef(const std::shared_ptr<sqlite3>& db)
        : _db(db)
      {}
      virtual std::shared_ptr<sqlite3> getDB() const
      {
        return _db;
      }
      std::shared_ptr<sqlite3> _db;
    };

    XLO_FUNC_START(xloSqlDB())
    {
      throwInFunctionWizard();

      return returnValue(
        cacheAdd(
          make_unique<DataBaseRef>(
            newDatabase())));
    }
    XLO_FUNC_END(xloSqlDB).threadsafe();
  }
}