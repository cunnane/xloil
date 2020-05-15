#include <sqlite/sqlite3ext.h>
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>
#include "Common.h"

#include <xlOil/Preprocessor.h>

using std::shared_ptr;
using std::vector;
using std::wstring;

namespace xloil
{
  namespace SQL
  {
    void processMeta(
      const ExcelArray& metaData, 
      sqlite3* db,
      ExcelObj::row_t i, 
      const ExcelObj& arg, 
      const wchar_t* defaultName)
    {
      if (!arg.isNonEmpty())
        return;

      if (i >= metaData.nRows() || metaData.nCols() < 1 || metaData(i, 0).isMissing())
        createVTable(db, ExcelArray(arg), defaultName);
      else
      { 
        vector<wstring> headings;
        std::transform(
          metaData.row_begin(i) + 1, metaData.row_end(i),
          std::back_inserter(headings),
          [](const ExcelObj& x) { return x.toString(); });
        createVTable(
          db,
          ExcelArray(arg),
          metaData(i, 0).toString().c_str(),
          headings.empty() ? nullptr : &headings);
      }
    }
  
#define XLOSQL_ARG_NAME Table
#define XLOSQL_NARGS 10

    constexpr wchar_t* TABLE_ARG_HELP = L"An array of data with rows as records";

    XLO_FUNC_START(
      xloSql(
        const ExcelObj& query,
        const ExcelObj& meta,
        XLO_DECLARE_ARGS(XLOSQL_NARGS, XLOSQL_ARG_NAME)
      )
    )
    {
      auto db = newDatabase();

      if (meta.isNonEmpty())
      {
        ExcelArray metaData(meta);
        ProcessArgs([db, metaData](auto iArg, auto& argVal, auto& argName)
        {
          processMeta(metaData, db.get(), iArg, argVal, argName);
        }, XLO_ARGS_LIST(XLOSQL_NARGS, XLOSQL_ARG_NAME));
      }
      else
      {
        ProcessArgs([db](auto& argVal, auto& argName)
        {
          if (argVal.isNonEmpty())
            createVTable(db.get(), ExcelArray(argVal), argName);
        }, XLO_ARGS_LIST(XLOSQL_NARGS, XLOSQL_ARG_NAME));
      }

      auto sql = query.toString();

      auto stmt = sqlPrepare(db.get(), sql);

      return ExcelObj::returnValue(sqlQueryToArray(stmt));
    }
    XLO_FUNC_END(xloSql).threadsafe()
      .help(L"Excecutes the SQL query on the provided tables. "
        "The tables will be named Table1, Table2, unless overrided "
        "by the Meta.  In the Meta, the first column should contain "
        "the names of the tables. Subsequent columns are interpreted "
        "as column headings for the table.")
      .arg(L"Query", L"The SQL query to perform")
      .arg(L"Meta", L"[opt] an array giving table names and column names")
      XLO_WRITE_ARG_HELP(XLOSQL_NARGS, XLOSQL_ARG_NAME, TABLE_ARG_HELP);
  }
}