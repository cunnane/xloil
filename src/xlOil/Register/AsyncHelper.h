#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Register.h>
#include <xlOil/ExcelCall.h>
#include <functional>

namespace xloil
{
  XLOIL_EXPORT void asyncReturn(const ExcelObj& asyncHandle, const ExcelObj& value);

  XLOIL_EXPORT bool yieldAndCheckIfEscPressed();

  XLOIL_EXPORT size_t lastCalcCancelledTicks();

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
      const ExcelObj* result = _call();
      asyncReturn(_asyncHandle, *result);
      if (result->xltype & msxll::xlbitDLLFree)
        delete result;
    }
  private:
    std::function<ExcelObj*()> _call;
    ExcelObj _asyncHandle;
  };
}