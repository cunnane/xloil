#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Register.h>
#include <xlOil/ExcelCall.h>
#include <functional>

namespace xloil
{
  class AsyncHolder
  {
  public:
    // No need to copy the data as FuncRegistry will keep this alive
    // Async handle is destroyed by Excel return, so must copy that
    AsyncHolder(std::function<ExcelObj*()> func, const ExcelObj* asyncHandle)
      : _call(func)
      , _asyncHandle(*asyncHandle)
    {
    }
    void operator()(int /*threadId*/) const
    {
      const ExcelObj* callBackArgs[2];
      callBackArgs[0] = &_asyncHandle;
      callBackArgs[1] = _call();
      // Need to use a raw call as the return value from xlAsyncReturn seems 
      // to be garbage - just a zeroed block of memory
      ExcelObj result;
      callExcelRaw(msxll::xlAsyncReturn, &result, 2, callBackArgs);
      if (callBackArgs[1]->xltype & msxll::xlbitDLLFree)
        delete callBackArgs[1];
    }
  private:
    std::function<ExcelObj*()> _call;
    ExcelObj _asyncHandle;
  };

  XLOIL_EXPORT bool yieldAndCheckIfEscPressed();

  XLOIL_EXPORT size_t lastCalcCancelledTicks();
}