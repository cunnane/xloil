#include <xloil/ExcelObjCache.h>
#include <xloil/ObjectCache.h>
#include <xloil/ExcelObj.h>
#include <xloil/StaticRegister.h>

using std::make_shared;
using std::make_unique;
using std::shared_ptr;
using std::unique_ptr;

namespace xloil
{
  static ObjectCache<unique_ptr<const ExcelObj>, detail::theObjectCacheUnquifier, false> 
    theExcelObjCache;
  
  XLOIL_EXPORT ExcelObj objectCacheAdd(unique_ptr<const ExcelObj>&& obj)
  {
    return theExcelObjCache.add(std::forward<unique_ptr<const ExcelObj>>(obj));
  }
  XLOIL_EXPORT bool objectCacheFetch(
    const std::wstring_view& cacheString, const ExcelObj*& obj)
  {
    unique_ptr<const ExcelObj>* cacheObj = nullptr;
    auto ret = theExcelObjCache.fetch(cacheString, cacheObj);
    obj = cacheObj->get();
    return ret;
  }
}

using namespace xloil;

XLO_FUNC_START(
  xloRef(const ExcelObj& pxOper)
)
{
  return returnValue(
    theExcelObjCache.add(
      make_unique<const ExcelObj>(pxOper)));
}
XLO_FUNC_END(xloRef).threadsafe()
  .help(L"Adds the specified value or range or array to the object cache and "
         "returns a string reference")
  .arg(L"ValOrArray", L"Data to be stored");


XLO_FUNC_START(
  xloVal(const ExcelObj& pxOper)
)
{
  unique_ptr<const ExcelObj>* result;
  // We return a pointer to the stored object directly without setting
  // the flag which tells Excel to free it.
  if (theExcelObjCache.fetch(pxOper.asPascalStr().view(), result))
    return returnReference(**result);

  return returnValue(CellError::Value);
}
XLO_FUNC_END(xloVal).threadsafe()
  .help(L"Given a string reference, returns a stored array or value. Cached values "
         "are not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) "
         "on workbook open")
  .arg(L"CacheRef", L"Cache reference string");

