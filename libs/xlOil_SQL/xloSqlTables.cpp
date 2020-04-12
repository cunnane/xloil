
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
    XLO_FUNC_START( xloSqlTables(
      const ExcelObj& database)
    )
    {
      Core::throwInFunctionWizard();

      std::shared_ptr<const CacheObj> dbObj;
      if (!cacheFetch(database.toString(), dbObj) || !dbObj)
        XLO_THROW("No database provided");
        
      auto stmt = sqlPrepare(dbObj->getDB().get(), 
        L"SELECT name FROM sqlite_master"
        "WHERE type = 'table' AND name NOT LIKE 'sqlite_%'");

      return ExcelObj::returnValue(sqlQueryToArray(stmt));
    }
    XLO_FUNC_END(xloSqlTables).threadsafe();
  }
}