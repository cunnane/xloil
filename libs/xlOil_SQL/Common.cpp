#include "Common.h"
#include <xloil/Interface.h>
#include "XlArrayTable.h"
#include <xlOil/ExcelArray.h>

using std::shared_ptr;
using std::string;
using std::wstring;
using std::vector;
using std::unique_ptr;

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
      if (sqlite3_open(":memory:", &db) == SQLITE_OK
        && sqlite3_create_module(db, "xlarray", &XlArrayModule, 0) == SQLITE_OK)
        return shared_ptr<sqlite3>(db, sqlite3_close);

      string msg(sqlite3_errmsg(db));
      sqlite3_close(db);
      XLO_THROW(msg);
    }

    wstring tableSchema(
      const ExcelArray& arr,
      const vector<wstring>* headings)
    {
      auto nCols = arr.nCols();
      wstring sql;
      sql.reserve(20 + nCols * 14);
      sql += L"CREATE TABLE x(";
      if (headings && headings->size() != arr.nCols())
        XLO_THROW("Provided {0} headings, but data has {1} columns", headings->size(), arr.nCols());

      for (auto j = 0u; j < nCols; ++j)
      {
        sql += headings
          ? headings->at(j)
          : arr(0, j).toString();
        auto col = arr.subArray(headings ? 0 : 1, j, -1, j + 1);
        auto colType = col.dataType();
        switch (colType)
        {
        case ExcelType::Int:
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

    void createVTable(
      sqlite3* db,
      const ExcelArray& arr,
      const wchar_t* name,
      const vector<wstring>* headings)
    {
      auto schema = tableSchema(arr, headings);
      auto arrayData = arr.subArray(headings ? 0 : 1, 0);
      auto sql = fmt::format(
        L"CREATE VIRTUAL TABLE {0} USING xlarray({1},{2},0)", name, (long long)&arrayData, schema);

      sqlite3_stmt *stmt;
      if (sqlite3_prepare16_v2(db, sql.c_str(),
        (int)sql.length() * sizeof(wchar_t), &stmt, 0) == SQLITE_OK)
      {
        sqlite3_step(stmt);
        sqlite3_finalize(stmt);
      }
      else
        XLO_THROW(L"Failed to create virtual table {0}", name);
    }

    shared_ptr<sqlite3_stmt> sqlPrepare(sqlite3* db, const wstring& sql)
    {
      sqlite3_stmt *stmt;
      sqlThrow(db,
        sqlite3_prepare16_v2(db, sql.c_str(), (int)sql.length() * sizeof(wchar_t), &stmt, 0));
      shared_ptr<sqlite3_stmt> stmtPtr(stmt, sqlite3_finalize);
      return stmtPtr;
    }

    int sqlExec(sqlite3* db, const wstring& sql)
    {
      const wchar_t* pSql = sql.c_str();
      const wchar_t* pSqlEnd = pSql + sql.length();

      int rc = SQLITE_OK;
      while (rc == SQLITE_OK && pSql != pSqlEnd)
      {
        int nCol = 0;

        const wchar_t* pLeftover;
        sqlite3_stmt *stmt;
        
        rc = sqlite3_prepare16_v2(db, pSql,
          (int)(pSqlEnd - pSql) * sizeof(wchar_t), &stmt, (const void **)&pLeftover);

        if (!stmt)
        {
          /* this happens for a comment or white-space */
          pSql = pLeftover;
          continue;
        }
        while (sqlite3_step(stmt) == SQLITE_ROW);
        rc = sqlite3_finalize(stmt);
        pSql = pLeftover;
        while (isspace(pSql[0])) pSql++;
      }

      return rc;
    }

    ExcelObj sqlQueryToArray(const std::shared_ptr<sqlite3_stmt>& prepared)
    {
      // Since we don't know the number of results in advance, our strategy is
      // to collect ExcelObj values in a vector and (pascal) strings in separate 
      // store. When we see a string, we put an empty string ExcelObj in results 
      // (this allocates no internal storage) and the string value in strings.
      // Later we allocate a contiguous block for both of these this data, copy 
      // them and fix-up the string objects to point to the right place.
      vector<ExcelObj> results;
      vector<wchar_t> strings;

      auto rc = sqlite3_step(prepared.get());
      auto nCols = sqlite3_column_count(prepared.get());
      int nRows = 0;
      while (rc == SQLITE_ROW)
      {
        for (auto j = 0; j < nCols; ++j)
        {
          switch (sqlite3_column_type(prepared.get(), j))
          {
          case SQLITE_INTEGER:
            results.emplace_back(sqlite3_column_int(prepared.get(), j));
            break;
          case SQLITE_FLOAT:
            results.emplace_back(sqlite3_column_double(prepared.get(), j));
            break;
          case SQLITE_TEXT:
          {
            auto text = (const wchar_t*)sqlite3_column_text16(prepared.get(), j);
            auto len = std::min<wchar_t>(USHRT_MAX, (unsigned short)wcslen(text));
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
        rc = sqlite3_step(prepared.get());
        ++nRows;
      }

      if (results.empty())
        return Const::Error(CellError::NA);

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

      return ExcelObj((const ExcelObj*)arrayData, nRows, nCols);
    }
  }
}