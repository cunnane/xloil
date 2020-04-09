#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include "ExcelArray.h"
#include "Common.h"
#include "Cache.h"

using std::shared_ptr;
using std::vector;
using std::make_shared;
using std::wstring;

namespace xloil
{
  namespace SQL
  {
    XLO_FUNC xloSqlTable(
      const ExcelObj& database,
      const ExcelObj& data,
      const ExcelObj& name,
      const ExcelObj& headings,
      const ExcelObj& query)
    {
      try
      {
        Core::throwInFunctionWizard();

        vector<wstring> headingsVec;
        if (headings.isNonEmpty())
        {
          ExcelArray headingArray(headings);
          std::transform(
            headingArray.begin(), headingArray.end(),
            std::back_inserter(headingsVec),
            [](const ExcelObj& x) { return x.toString(); });
        }

        std::shared_ptr<const CacheObj> dbObj;
        if (!cacheFetch(database.toString(), dbObj))
          XLO_THROW("No database provided");
        
        auto tableName = name.toString();

        auto db = dbObj->getDB();
        ScopedLock lock(db.get());

        createVTable(
          db.get(),
          ExcelArray(data),
          tableName.c_str(),
          headingsVec.empty() ? nullptr : &headingsVec);

        wstring select = query.isNonEmpty()
          ? query.toString()
          : fmt::format(L"SELECT * FROM {0}", tableName);

        auto tempName = wstring(L"xloil_temp");
        auto sql = fmt::format(
          L"CREATE TABLE {0} AS {1};"
          "DROP TABLE {2};"
          "ALTER TABLE {0} RENAME TO {2};",
          tempName, select, tableName);
        sqlExec(db.get(), sql);
        
        return const_cast<ExcelObj*>(&database);
      }
      catch (const std::exception& e)
      {
        XLO_RETURN_ERROR(e);
      }
    }
    XLO_REGISTER(xloSqlTable).threadsafe()
      .help(L"Creates a table in a database optionally via a query (otherwise uses SELECT *). "
            "Returns a reference to the database: it is recommended to chain xloSqlTable calls "
            "to clarify execution order")
      .arg(L"Database")
      .arg(L"Data")
      .arg(L"Name")
      .arg(L"Headings")
      .arg(L"Query", L"[opt] A select statement used to extract a subset of the data");
  }
}