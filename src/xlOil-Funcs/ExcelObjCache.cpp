#include <xloil/ExcelObjCache.h>
#include <xloil/ObjectCache.h>
#include <xloil/ExcelObj.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  decltype(ObjectCacheFactory<std::unique_ptr<const ExcelObj>>::cache) ObjectCacheFactory<std::unique_ptr<const ExcelObj>>::cache;
}

using namespace xloil;

XLO_FUNC_START(
  xloRef(const ExcelObj& pxOper)
)
{
  return returnValue(makeCached<ExcelObj>(pxOper));
}
XLO_FUNC_END(xloRef).threadsafe()
  .help(L"Adds the specified value or range or array to the object cache and "
         "returns a string reference")
  .arg(L"ValOrArray", L"Data to be stored");


XLO_FUNC_START(
  xloVal(const ExcelObj& pxOper)
)
{
  // We return a pointer to the stored object directly without setting
  // the flag which tells Excel to free it.
  auto result = getCached<ExcelObj>(pxOper.asPString().view());
  if (result)
    return returnReference(*result);
  return returnValue(CellError::Value);
}
XLO_FUNC_END(xloVal).threadsafe()
  .help(L"Given a string reference, returns a stored array or value. Cached values "
         "are not saved so will need to be recreated by a full recalc (Ctrl-Alt-F9) "
         "on workbook open")
  .arg(L"CacheRef", L"Cache reference string");

