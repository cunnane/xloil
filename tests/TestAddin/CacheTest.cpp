#include <xloil/ExcelObjCache.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START(
    testCacheOut(const ExcelObj& inArray)
  )
  {
    auto key = makeCached<int>(7);
    return returnValue(key);
  }
  XLO_FUNC_END(testCacheOut);
}