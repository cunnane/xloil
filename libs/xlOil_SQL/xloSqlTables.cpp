
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
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
    XLO_FUNC xloSqlTables(
      const ExcelObj& database)
    {
      try
      {
        if (Core::inFunctionWizard())
          XLO_THROW("In wizard");

        std::shared_ptr<const CacheObj> dbObj;
        if (!cacheFetch(database.toString(), dbObj) || !dbObj)
          XLO_THROW("No database provided");
        
        auto stmt = sqlPrepare(dbObj->getDB().get(), 
          L"SELECT name FROM sqlite_master"
          "WHERE type = 'table' AND name NOT LIKE 'sqlite_%'");

        return ExcelObj::returnValue(sqlQueryToArray(stmt));
      }
      catch (const std::exception& e)
      {
        XLO_RETURN_ERROR(e);
      }
    }
    XLO_REGISTER(xloSqlTables).threadsafe();
  }
}