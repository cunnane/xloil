#define XLOIL_UNSAFE_INPLACE_RETURN

#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/Async.h>
#include <xloil/FPArray.h>
#include <xloil/ExcelRef.h>

using std::shared_ptr;

namespace xloil
{
  XLO_ENTRY_POINT(void) testInPlace(
    ExcelObj* retVal,
    ExcelObj* nRows,
    const ExcelObj* nCols)
  {
    try
    {
      if (retVal->isType(ExcelType::Multi))
      {
        auto& arr = retVal->val.array;
        arr.rows = std::min(nRows->toInt(1), arr.rows);
        arr.columns = std::min(nCols->toInt(1), arr.columns);
      }
    }
    catch (const std::exception& e)
    {
      *retVal = *returnValue(e);
    }
  }
  XLO_REGISTER_FUNC(testInPlace).threadsafe();

  XLO_ENTRY_POINT(void) testAsync(
    const AsyncHandle& handle,
    const ExcelObj& val
  )
  {
    try
    {
      handle.returnValue(val);
    }
    catch (const std::exception& e)
    {
      handle.returnValue(e.what());
    }
  }
  XLO_REGISTER_FUNC(testAsync);

  XLO_FUNC_START(testFP(const FPArray& array))
  {
    ExcelArrayBuilder builder(array.rows, array.columns);
    for (auto i = 0; i < array.rows; ++i)
      for (auto j = 0; j < array.columns; ++j)
        builder(j, i) = array(i, j);
    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(testFP);

  XLO_FUNC_START(testRangeAddy(const RangeArg& range))
  {
    return returnValue(range.address());
  }
  XLO_FUNC_END(testRangeAddy);

  XLO_ENTRY_POINT(void) testTranspose(
    FPArray& array
  )
  {
    for (auto i = 0; i < array.rows; ++i)
      for (auto j = i + 1; j < array.columns; ++j)
        std::swap(array(j, i), array(i, j));
  }
  XLO_REGISTER_FUNC(testTranspose).threadsafe();

  XLO_ENTRY_POINT(int) testCommand(
    const ExcelObj*, const ExcelObj*
  )
  {
    return 1;
  }
  XLO_REGISTER_FUNC(testCommand).command();
}
