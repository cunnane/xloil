#include <xloil/ExcelObjCache.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START(
    testCacheIn(const ExcelObj& cacheKey)
  )
  {
    auto val = getCached<int>(cacheKey.asPString());
    return returnValue(val ? *val : 0);
  }
  XLO_FUNC_END(testCacheIn);
}