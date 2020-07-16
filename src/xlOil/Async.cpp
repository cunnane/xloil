#include <xloil/Async.h>
#include <xlOil/Events.h>
#include <xloil/ExcelCall.h>
#include <xlOilHelpers/WindowsSlim.h>

namespace xloil
{
  XLOIL_EXPORT void asyncReturn(
    const ExcelObj& asyncHandle, const ExcelObj& value)
  {
    const ExcelObj* callBackArgs[2];
    callBackArgs[0] = &asyncHandle;
    callBackArgs[1] = &value;
    // Need to use a raw call as the return value from xlAsyncReturn seems 
    // to be garbage - just a zeroed block of memory
    ExcelObj result;
    callExcelRaw(msxll::xlAsyncReturn, &result, 2, callBackArgs);
  }

  XLOIL_EXPORT bool yieldAndCheckIfEscPressed()
  {
    auto[res, ret] = tryCallExcel(msxll::xlAbort);
    return (ret == 0 && res.asBool());
  }
}

