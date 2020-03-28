#include <sqlite/sqlite3.h>
#include <sqlite/sqlite3ext.h>
#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include "ArrayBuilder.h"
#include "ExcelArray.h"
#include "XlArrayTable.h"

using std::shared_ptr;
using std::string;
using std::wstring;
using std::vector;
using std::unique_ptr;

#define XLO_WSTR(x) L ## #x

namespace xloil
{
  namespace SQL
  {
    void sqlThrow(sqlite3* db, int errCode)
    {
      if (errCode != SQLITE_OK)
        XLO_THROW(sqlite3_errmsg(db));
    }

    shared_ptr<sqlite3> newDatabase()
    {
      sqlite3 *db;
      if ( sqlite3_open(":memory:", &db) == SQLITE_OK 
        && sqlite3_create_module(db, "xlarray", &XlArrayModule, 0) == SQLITE_OK)
        return shared_ptr<sqlite3>(db, sqlite3_close);

      string msg(sqlite3_errmsg(db));
      sqlite3_close(db);
      XLO_THROW(msg);
    }

    wstring tableSchema(const ExcelArray& arr)
    {
      auto nCols = arr.nCols();
      wstring sql(L"CREATE TABLE x(");
      for (auto j = 0; j < nCols; ++j)
      {
        auto heading = arr(0, j).toString();
        auto col = arr.subArray(1, j, -1, j + 1);
        sql += heading;
        auto colType = col.dataType();
        switch (colType)
        {
        case ExcelType::Int:
          sql += L" INTEGER";
          break;
        case ExcelType::Bool:
          sql += L" INTEGER";
          break;
        case ExcelType::Num:
          sql += L" REAL";
          break;
        case ExcelType::Str:
          sql += L" TEXT";
          break;
        default:
          break;
        }
        sql += L',';
      }
      if (nCols > 0)
        sql.pop_back();
      sql += L')';

      return sql;
    }

    void createVTable(sqlite3* db, const ExcelArray& arr, const wchar_t* name)
    {
      auto schema = tableSchema(arr);
      auto arrayData = arr.subArray(1, 0);
      auto sql = fmt::format(
        L"CREATE VIRTUAL TABLE {0} USING xlarray({1},{2})", name, (long long)&arrayData, schema);

      sqlite3_stmt *stmt;
      auto rc = sqlite3_prepare16_v2(db, sql.c_str(), sql.length() * sizeof(wchar_t), &stmt, 0);
      sqlite3_step(stmt);
      sqlite3_finalize(stmt);
    }

    XLO_FUNC xloSql(
      const ExcelObj& query, 
      const ExcelObj& table1, 
      const ExcelObj& table2,
      const ExcelObj& table3,
      const ExcelObj& table4,
      const ExcelObj& table5)
    {
      try
      {
        auto db = newDatabase();

#define XLO_CREATE_VTAB(tab) \
  if (tab.isNonEmpty()) createVTable(db.get(), ExcelArray(tab), XLO_WSTR(tab))
        
        XLO_CREATE_VTAB(table1);
        XLO_CREATE_VTAB(table2);
        XLO_CREATE_VTAB(table3);
        XLO_CREATE_VTAB(table4);
        XLO_CREATE_VTAB(table5);


        auto sql = query.toString();

        sqlite3_stmt *stmt;
        sqlThrow(db.get(), 
          sqlite3_prepare16_v2(db.get(), sql.c_str(), sql.length() * sizeof(wchar_t), &stmt, 0));
        shared_ptr<sqlite3_stmt> stmtPtr(stmt, sqlite3_finalize);

        /* Iterate over results */
        auto nCols = sqlite3_column_count(stmt);
        auto rc = sqlite3_step(stmt);

        // Since we don't know the number of results in advance, our strategy is
        // to collect ExcelObj values in a vector and (pascal) strings in separate 
        // store. When we see a string, we put an empty string ExcelObj in results 
        // (this allocates no internal storage) and the string value in strings.
        // Later we allocate a contiguous block for both of these this data, copy 
        // them and fix-up the string objects to point to the right place.
        vector<ExcelObj> results;
        vector<wchar_t> strings;

        int nRows = 0;
        while (rc == SQLITE_ROW)
        {
          for (auto j = 0; j < nCols; ++j)
          {
            switch (sqlite3_column_type(stmt, j))
            {
            case SQLITE_INTEGER:
              results.emplace_back(sqlite3_column_int(stmt, j));
              break;
            case SQLITE_FLOAT:
              results.emplace_back(sqlite3_column_double(stmt, j));
              break;
            case SQLITE_TEXT:
            {
              auto text = (const wchar_t*)sqlite3_column_text16(stmt, j);
              auto len = std::min<wchar_t>(USHRT_MAX, wcslen(text));
              strings.push_back(len);
              strings.insert(strings.end(), text, text + len);
              // Empty string into results - we will fix it later
              results.emplace_back(ExcelType::Str);
              break;
            }
            case SQLITE_BLOB:
            case SQLITE_NULL:
            default:
              results.emplace_back(CellError::NA);
              break;
            }
          }
          rc = sqlite3_step(stmt);
          ++nRows;
        }

        if (results.empty())
          return ExcelObj::returnValue(Const::Error(CellError::NA));

        // Memcpy the ExcelObj and string values into a single block
        // Since Excel expects arrays by-row, we don't need to re-order
        auto resultsNBytes = sizeof(ExcelObj) * results.size();
        auto stringsNBytes = sizeof(wchar_t) * strings.size();
        auto* arrayData = new char[resultsNBytes + stringsNBytes];
        auto* stringData = arrayData + resultsNBytes;
        memcpy_s(arrayData, resultsNBytes, results.data(), resultsNBytes);
        memcpy_s(stringData, stringsNBytes, strings.data(), stringsNBytes);

        // Fix up strings
        auto pObj = (ExcelObj*)arrayData;
        auto pEnd = pObj + results.size();
        auto pStr = (wchar_t*)stringData;
        for (; pObj != pEnd; ++pObj)
        {
          if (pObj->xltype == msxll::xltypeStr)
          {
            assert((char*)pStr < stringData + stringsNBytes);
            pObj->val.str = pStr;
            // Skip to the next string by adding the string length
            pStr += pStr[0] + 1;
          }
        }
        
        return ExcelObj::returnValue((const ExcelObj*)arrayData, nRows, nCols);
      }
      catch (const std::exception& e)
      {
        XLO_RETURN_ERROR(e);
      }
    }
    XLO_REGISTER(xloSql)
      .help(L"Given a string reference, returns a stored array or cell value. "
        "The cache is not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) on workbook open")
      .arg(L"CacheRef", L"Cache reference string")
      .threadsafe();
  }

  Core* theCore = nullptr;

  XLO_PLUGIN_INIT(Core& core)
  {
    theCore = &core;
    spdlog::set_default_logger(core.getLogger());
    return 0;
  }
}