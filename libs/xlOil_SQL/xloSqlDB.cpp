
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include <xloil/ExcelState.h>
#include "ExcelArray.h"
#include "Common.h"
#include "Cache.h"

using std::shared_ptr;
using std::vector;
using std::make_shared;

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

    XLO_FUNC xloSqlDB()
    {
      try
      {
        if (Core::inFunctionWizard())
          XLO_THROW("In wizard");

        return ExcelObj::returnValue(
          cacheAdd(
            make_shared<DataBaseRef>(
              newDatabase())));
      }
      catch (const std::exception& e)
      {
        XLO_RETURN_ERROR(e);
      }
    }
    XLO_REGISTER(xloSqlDB).threadsafe();
  }
}