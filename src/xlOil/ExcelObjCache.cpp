#include "ExcelObjCache.h"
#include "ObjectCache.h"
#include "ExcelObj.h"
#include "Register.h"

using std::make_shared;
using std::shared_ptr;

namespace xloil
{
  static ObjectCache<std::shared_ptr<const ExcelObj>, false> theExcelObjCache 
    = ObjectCache<std::shared_ptr<const ExcelObj>, false>(theObjectCacheUnquifier);
  
  ExcelObj addCacheObject(const shared_ptr<const ExcelObj>& obj)
  {
    return theExcelObjCache.add(obj);
  }
  bool fetchCacheObject(const wchar_t* cacheString, size_t length, shared_ptr<const ExcelObj>& obj)
  {
    return theExcelObjCache.fetch(cacheString, length, obj);
  }
}

using namespace xloil;

XLO_FUNC xloRef(ExcelObj* pxOper)
{
  try
  {
    auto q = make_shared<const ExcelObj>(*pxOper);
    return ExcelObj::returnValue(theExcelObjCache.add(q));
  }
  catch (...)
  {
    return new ExcelObj();
  }
}
XLO_REGISTER(xloRef)
.help(L"Adds the specified cell or array to the object cache and returns a string reference")
.arg(L"CellOrArray", L"Data to be stored");


XLO_FUNC xloDeref(ExcelObj* pxOper)
{
  try
  {
    shared_ptr<const ExcelObj> result;
    // TODO: can we do the regex match without this string copy?
    // TODO: check the followng is safe

    // We return a pointer to the stored object directly without setting
    // the flag which tells Excel to free it.
    if (theExcelObjCache.fetch(pxOper->toString().c_str(), result))
      return const_cast<ExcelObj*>(result.get());

    return ExcelObj::returnValue(CellError::NA);
  }
  catch (...)
  {
    return new ExcelObj();
  }
}
XLO_REGISTER(xloDeref)
.help(L"Given a string reference, returns a stored array or cell value. "
  "The cache is not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) on workbook open")
.arg(L"CacheRef", L"Cache reference string");
