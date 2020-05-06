#include "AsyncHelper.h"
#include <xlOil/Events.h>
#include <xlOilHelpers/WindowsSlim.h>

namespace xloil
{
  static size_t lastCancelTime = 0;

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

  XLOIL_EXPORT size_t lastCalcCancelledTicks()
  {
    return lastCancelTime;
  }

  namespace
  {
    struct RegisterMe
    {
      RegisterMe()
      {
        static auto handler = xloil::Event_CalcCancelled() += []()
        { 
#ifdef _WIN64
          lastCancelTime = GetTickCount64(); 
#else
          lastCancelTime = GetTickCount();
#endif
        };
      }
    } theInstance;
  }
}

