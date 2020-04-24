#pragma once
#define SQLITE_OMIT_PROGRESS_CALLBACK
#define SQLITE_OMIT_AUTHORIZATION
#define SQLITE_OMIT_WAL

#include <xlOil/ExcelObj.h>
#include <sqlite/sqlite3.h>
#include <sqlite/sqlite3ext.h>
#include <memory>
#include <string>
#include <vector>

namespace xloil { class ExcelArray; }

namespace xloil
{
  namespace SQL
  {
    void 
      sqlThrow(sqlite3* db, int errCode);

    std::shared_ptr<sqlite3> 
      newDatabase();

    void 
      createVTable(
        sqlite3* db, 
        const ExcelArray& arr, 
        const wchar_t* name, 
        const std::vector<std::wstring>* headings = nullptr);

    std::shared_ptr<sqlite3_stmt> 
      sqlPrepare(sqlite3* db, const std::wstring& sql);

    int 
      sqlExec(sqlite3* db, const std::wstring& sql);

    ExcelObj
      sqlQueryToArray(const std::shared_ptr<sqlite3_stmt>& prepared);

    class ScopedLock
    {
    public:
      ScopedLock(sqlite3* db) : _db(db)
      {
        sqlite3_mutex_enter(sqlite3_db_mutex(db));
      }
      ~ScopedLock()
      {
        sqlite3_mutex_leave(sqlite3_db_mutex(_db));
      }
    private:
      sqlite3* _db;
    };
  }
}
