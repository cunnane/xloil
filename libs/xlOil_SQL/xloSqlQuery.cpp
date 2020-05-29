
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include <xlOil/ExcelArray.h>
#include "Common.h"
#include "Cache.h"

using std::shared_ptr;
using std::vector;
using std::make_shared;

namespace xloil
{
  namespace SQL
  {
    XLO_FUNC_START( xloSqlQuery(
      const ExcelObj& database,
      const ExcelObj& query
      )
    )
    {
      if (Core::inFunctionWizard())
        XLO_THROW("In wizard");

      shared_ptr<const CacheObj> dbObj;
      if (!cacheFetch(database.toString(), dbObj))
        XLO_THROW("No database provided");

      auto sql = query.toString();
      auto stmt = sqlPrepare(dbObj->getDB().get(), sql);

      return returnValue(sqlQueryToArray(stmt));
    }
    XLO_FUNC_END(xloSqlQuery).threadsafe()
      .help(L"Runs the specified query on a database, returning the results as an array")
      .arg(L"Database", L"A cache reference to a database object created wth xloSqlDB")
      .arg(L"Query");
  }
}