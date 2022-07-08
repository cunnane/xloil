#include "XlArrayTable.h"
#include <sqlite/sqlite3.h>
#include <sqlite/sqlite3ext.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelRef.h>

using std::shared_ptr;
using std::pair;

namespace xloil
{
  namespace SQL
  {
    /* An instance of the XlArray virtual table */
   
    struct XlArrayTable
    {
      using InputType = XlArrayInput;
      XlArrayTable(const InputType& input) : data(input) {};
      sqlite3_vtab base;              /* Base class.  Must be first */
      ExcelArray data;
    };

    struct XlRangeTable
    {
      using InputType = XlRangeInput;
      XlRangeTable(const InputType& input) : data(input) {};
      sqlite3_vtab base;              /* Base class.  Must be first */
      ExcelRef data;
    };

    // Not sure what this could be used for...
    //
    //using XlArrayOwnerInput = std::pair<ExcelArray, std::shared_ptr<const ExcelObj>>;
    //struct XlArrayOwnerTable
    //{
    //  using InputType = XlArrayOwnerInput;
    //  XlArrayOwnerTable(const InputType& input)
    //   : data(input.first)
    //   , storage(input.second)
    //  {}
    //  sqlite3_vtab base;              /* Base class.  Must be first */
    //  ExcelArray data;
    //  shared_ptr<const ExcelObj> storage;
    //};


    struct XlTableCursor 
    {
      sqlite3_vtab_cursor base;  /* Base class.  Must be first */
      int iRowid;                /* The current rowid.  Negative for EOF */
    };

    template<class T>
    static int xConnect(
      sqlite3 *db,
      void* /*pAux*/,
      int /*argc*/, const char *const* argv,
      sqlite3_vtab **ppVtab,
      char** /*pzErr*/)
    {
      auto input = (const typename T::InputType*)atoll(argv[3]);
      auto schema = argv[4];

      auto rc = sqlite3_declare_vtab(db, schema);

      if (rc == SQLITE_OK)
      {
        auto* pNew = new T(*input);
        *ppVtab = (sqlite3_vtab*)pNew;
      }

      return rc;
    }

    /*
    ** The xConnect and xCreate methods do the same thing, but they must be
    ** different so that the virtual table is not an eponymous virtual table.
    */
    template<class T>
    static int xCreate(
      sqlite3 *db,
      void *pAux,
      int argc, const char *const*argv,
      sqlite3_vtab **ppVtab,
      char **pzErr)
    {
      return xConnect<T>(db, pAux, argc, argv, ppVtab, pzErr);
    }

    static int xBestIndex(
      sqlite3_vtab*,
      sqlite3_index_info *pIdxInfo)
    {
      pIdxInfo->estimatedCost = 1000000;
      return SQLITE_OK;
    }

    /*
    ** This is the destructor for the vtable.
    */
    template<class T>
    static int xDisconnect(sqlite3_vtab *pVtab) {
      auto* p = (T*)pVtab;
      delete p;
      return SQLITE_OK;
    }

    static int xOpen(sqlite3_vtab*, sqlite3_vtab_cursor **ppCursor) {
      auto *pCur = new XlTableCursor();
      *ppCursor = &pCur->base;
      return SQLITE_OK;
    }

    /*
    ** Destructor for a Cursor.
    */
    static int xClose(sqlite3_vtab_cursor *cur) {
      auto *pCur = (XlTableCursor*)cur;
      delete pCur;
      return SQLITE_OK;
    }

    /*
    ** Only a full table scan is supported.  So xFilter simply rewinds to
    ** the beginning.
    */
    static int xFilter(
      sqlite3_vtab_cursor *pVtabCursor,
      int /*idxNum*/, const char* /*idxStr*/,
      int /*argc*/, sqlite3_value** /*argv*/)
    {
      auto *pCur = (XlTableCursor*)pVtabCursor;
      pCur->iRowid = 0;
      return SQLITE_OK;
    }

    /*
    ** Advance a Cursor to its next row of input.
    ** Set the EOF marker if we reach the end of input.
    */
    template<class T>
    static int xNext(sqlite3_vtab_cursor *cur)
    {
      auto *pCur = (XlTableCursor*)cur;
      auto *pTab = (const T*)cur->pVtab;
      if (++pCur->iRowid >= (int)pTab->data.nRows())
        pCur->iRowid = -1;
      return SQLITE_OK;
    }

    /*
    ** Return TRUE if the cursor has been moved off of the last
    ** row of output.
    */
    static int xEof(sqlite3_vtab_cursor *cur) {
      auto *pCur = (const XlTableCursor*)cur;
      return pCur->iRowid < 0;
    }

    
    template<class T> struct TableReturn {};
    template<> struct TableReturn<XlRangeTable> { using type = ExcelObj; };
    template<> struct TableReturn<XlArrayTable> { using type = const ExcelObj&; };


    struct ExcelValToSqlType
    {
      sqlite3_context* ctx;

      void operator()(int x) const     { sqlite3_result_int64(ctx, x); }
      void operator()(bool x) const    { sqlite3_result_int64(ctx, x); }
      void operator()(double x) const  { sqlite3_result_double(ctx, x); }
      void operator()(CellError) const { sqlite3_result_null(ctx); }
      void operator()(nullptr_t) const { sqlite3_result_null(ctx); }
      void operator()(const PStringRef& pstr) const 
      {
        sqlite3_result_text16(ctx, pstr.pstr(),
          pstr.length() * sizeof(wchar_t), SQLITE_STATIC);
      }

      template <class T> void operator()(T) const
      {
        sqlite3_result_error(ctx, "Unexpected excel type", -1);
      }
    };

    /*
    ** Return values of columns for the row at which the cursor
    ** is currently pointing.
    */
    template<class T>
    static int xColumn(
      sqlite3_vtab_cursor *cur,   /* The cursor */
      sqlite3_context *ctx,       /* First argument to sqlite3_result_...() */
      int i)                      /* Which column to return */
    {
      auto *pCur = (XlTableCursor*)cur;
      auto *pTab = (const T*)cur->pVtab;
      TableReturn<T>::type val = pTab->data(pCur->iRowid, i);
      val.visit(ExcelValToSqlType{ ctx });
      return SQLITE_OK;
    }

    /*
    ** Return the rowid for the current row.
    */
    static int xRowid(sqlite3_vtab_cursor *cur, sqlite_int64 *pRowid)
    {
      auto *pCur = (XlTableCursor*)cur;
      *pRowid = pCur->iRowid;
      return SQLITE_OK;
    }

    template<class T>
    sqlite3_module XlModule = {
      0,                  /* iVersion */
      xCreate<T>,         /* xCreate */
      xConnect<T>,        /* xConnect */
      xBestIndex,         /* xBestIndex */
      xDisconnect<T>,     /* xDisconnect */
      xDisconnect<T>,     /* xDestroy */
      xOpen,              /* xOpen - open a cursor */
      xClose,             /* xClose - close a cursor */
      xFilter,            /* xFilter - configure scan constraints */
      xNext<T>,           /* xNext - advance a cursor */
      xEof,               /* xEof - check for end of scan */
      xColumn<T>,         /* xColumn - read data */
      xRowid,             /* xRowid - read data */
      0,                  /* xUpdate */
      0,                  /* xBegin */
      0,                  /* xSync */
      0,                  /* xCommit */
      0,                  /* xRollback */
      0,                  /* xFindMethod */
      0,                  /* xRename */
    };


    extern sqlite3_module XlArrayModule = XlModule<XlArrayTable>;
    extern sqlite3_module XlRangeModule = XlModule<XlRangeTable>; 
  }
}