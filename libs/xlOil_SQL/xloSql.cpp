#include <sqlite/sqlite3ext.h>
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include "ExcelArray.h"
#include "ExcelObj.h"
#include "Common.h"
#include <boost/preprocessor/repeat_from_to.hpp>
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
      size_t i, 
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

    XLO_FUNC xloSql(
      const ExcelObj& query,
      const ExcelObj& meta,
      XLO_DECLARE_ARGS(XLOSQL_NARGS, XLOSQL_ARG_NAME)
    )
    {
      try
      {
        auto db = newDatabase();

        if (meta.isNonEmpty())
        {
          ExcelArray metaData(meta);
#define MY_CREATE_VTAB_META(z, N, prefix) \
          processMeta(metaData, db.get(), N - 1, prefix##N, XLO_WSTR(prefix##N));
          BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, XLOSQL_NARGS), MY_CREATE_VTAB_META, XLOSQL_ARG_NAME);
#undef MY_CREATE_VTAB_META
        }
        else
        {
#define MY_CREATE_VTAB(z, N, prefix) \
          if (prefix##N.isNonEmpty()) \
            createVTable(db.get(), ExcelArray(prefix##N), XLO_WSTR(prefix##N));

          BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, XLOSQL_NARGS), MY_CREATE_VTAB, XLOSQL_ARG_NAME);
#undef MY_CREATE_VTAB
        }

        auto sql = query.toString();

        auto stmt = sqlPrepare(db.get(), sql);

        return ExcelObj::returnValue(sqlQueryToArray(stmt));
      }
      catch (const std::exception& e)
      {
        XLO_RETURN_ERROR(e);
      }
    }

    constexpr wchar_t* TABLE_ARG_HELP = L"An array of data with rows as records";
#define WRITE_ARG_HELP(z, N, prefix) .arg(XLO_WSTR(prefix##N), TABLE_ARG_HELP)

    XLO_REGISTER(xloSql).threadsafe()
      .help(L"Excecutes the SQL query on the provided tables. "
        "The tables will be named Table1, Table2, unless overrided "
        "by the Meta.  In the Meta, the first column should contain "
        "the names of the tables. Subsequent columns are interpreted "
        "as column headings for the table.")
      .arg(L"Query", L"The SQL query to perform")
      .arg(L"Meta", L"[opt] an array giving table names and column names")
      BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, XLOSQL_NARGS), WRITE_ARG_HELP, XLOSQL_ARG_NAME);
  }
}