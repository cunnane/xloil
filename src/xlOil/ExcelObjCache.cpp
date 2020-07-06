#include <xloil/ExcelObjCache.h>
#include <xloil/ObjectCache.h>
#include <xloil/ExcelObj.h>
#include <xloil/StaticRegister.h>

using std::make_shared;
using std::shared_ptr;

namespace xloil
{
  // TODO: why a shared-ptr?
  static ObjectCache<shared_ptr<const ExcelObj>, detail::theObjectCacheUnquifier, false> theExcelObjCache;
  
  ExcelObj objectCacheAdd(shared_ptr<const ExcelObj>&& obj)
  {
    return theExcelObjCache.add(std::forward<shared_ptr<const ExcelObj>>(obj));
  }
  bool objectCacheFetch(const std::wstring_view& cacheString, shared_ptr<const ExcelObj>& obj)
  {
    return theExcelObjCache.fetch(cacheString, obj);
  }
}

using namespace xloil;

XLO_FUNC_START(
  xloRef(const ExcelObj& pxOper)
)
{
  return returnValue(
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
  // We return a pointer to the stored object directly without setting
  // the flag which tells Excel to free it.
  if (theExcelObjCache.fetch(pxOper.asPascalStr().view(), result))
    return returnReference(*result);

  return returnValue(CellError::Value);
}
XLO_FUNC_END(xloVal).threadsafe()
  .help(L"Given a string reference, returns a stored array or cell value. "
    "The cache is not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) on workbook open")
  .arg(L"CacheRef", L"Cache reference string");

