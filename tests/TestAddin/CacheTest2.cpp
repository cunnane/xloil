#include <xloil/ExcelObjCache.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START(
    testCacheIn(const ExcelObj& cacheKey)
  )
  {
    auto val = get_cached<int>(cacheKey.asPString());
    return returnValue(*val);
  }
  XLO_FUNC_END(testCacheIn);
}