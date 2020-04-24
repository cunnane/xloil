#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include <xlOil/ExcelArray.h>
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
    XLO_FUNC_START( xloSqlTable(
      const ExcelObj& database,
      const ExcelObj& data,
      const ExcelObj& name,
      const ExcelObj& headings,
      const ExcelObj& query)
    )
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

      // Attempt to drop table if it already exists, e.g. function called 
      // again, but ignore return code
      sqlExec(db.get(), fmt::format(L"DROP TABLE {0}", tableName));

      createVTable(
        db.get(),
        ExcelArray(data),
        tableName.c_str(),
        headingsVec.empty() ? nullptr : &headingsVec);

      wstring select = query.isNonEmpty()
        ? query.toString()
        : fmt::format(L"SELECT * FROM {0}", tableName);

      // We do this little rename so the table can have the 
      // expected name in the query even though it is just
      // the temporary vtable.
      auto tempName = wstring(L"xloil_temp");
      auto sql = fmt::format(
        L"CREATE TABLE {0} AS {1};"
        "DROP TABLE {2};"
        "ALTER TABLE {0} RENAME TO {2};",
        tempName, select, tableName);
      sqlThrow(db.get(), sqlExec(db.get(), sql));
        
      return const_cast<ExcelObj*>(&database);
    }
    XLO_FUNC_END(xloSqlTable).threadsafe()
      .help(L"Creates a table in a database, optionally via a query. "
            "Returns a reference to the database: it is recommended to chain xloSqlTable calls "
            "to force execution order before calling xloSqlQuery")
      .arg(L"Database", L"A reference to a database created by xloSqlDB or another call to xloSqlTable")
      .arg(L"Data", L"An array of data to read into the table. First row must be headings, unless "
                     "headings parameter is specified")
      .arg(L"Name", L"The table name in the database, this must be unique")
      .arg(L"Headings", L"[opt] headings (field names) for the data")
      .arg(L"Query", L"[opt] A select statement used to extract a subset of the data, otherwise SELECT * is used");
  }
}