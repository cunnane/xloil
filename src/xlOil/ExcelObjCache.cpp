#include "ExcelObjCache.h"
#include "ObjectCache.h"
#include "ExcelObj.h"
#include <xloil/StaticRegister.h>

using std::make_shared;
using std::shared_ptr;

namespace xloil
{
  // TODO: why a shared-ptr?
  static ObjectCache<shared_ptr<const ExcelObj>, theObjectCacheUnquifier, false> theExcelObjCache
    = ObjectCache<shared_ptr<const ExcelObj>, theObjectCacheUnquifier, false>();
  
  ExcelObj objectCacheAdd(shared_ptr<const ExcelObj>&& obj)
  {
    return theExcelObjCache.add(std::forward<shared_ptr<const ExcelObj>>(obj));
  }
  bool objectCacheFetch(const wchar_t* cacheString, size_t length, shared_ptr<const ExcelObj>& obj)
  {
    return theExcelObjCache.fetch(cacheString, length, obj);
  }
}

using namespace xloil;

XLO_FUNC_START(
  xloRef(const ExcelObj& pxOper)
)
{
  return ExcelObj::returnValue(
    theExcelObjCache.add(
      make_shared<const ExcelObj>(pxOper)));
}
XLO_FUNC_END(xloRef).threadsafe()
  .help(L"Adds the specified cell or array to the object cache and returns a string reference")
  .arg(L"CellOrArray", L"Data to be stored");


XLO_FUNC_START(
  xloVal(const ExcelObj& pxOper)
)
{
  shared_ptr<const ExcelObj> result;
  // TODO: can we do the regex match without this string copy?

  // We return a pointer to the stored object directly without setting
  // the flag which tells Excel to free it.
  if (theExcelObjCache.fetch(pxOper.toString().c_str(), result))
    return const_cast<ExcelObj*>(result.get());

  return ExcelObj::returnValue(CellError::Value);
}
XLO_FUNC_END(xloVal).threadsafe()
  .help(L"Given a string reference, returns a stored array or cell value. "
    "The cache is not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) on workbook open")
  .arg(L"CacheRef", L"Cache reference string");

