#include "AsyncHelper.h"
#include <xlOil/Events.h>
#include <xlOilHelpers/WindowsSlim.h>

namespace xloil
{
  static size_t lastCancelTime = 0;

  XLOIL_EXPORT  bool yieldAndCheckIfEscPressed()
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

