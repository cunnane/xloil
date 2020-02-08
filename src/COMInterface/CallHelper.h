#pragma once
#include "xloil/Log.h"

namespace xloil
{
  // See https://social.msdn.microsoft.com/Forums/vstudio/en-US/9168f9f2-e5bc-4535-8d7d-4e374ab8ff09/hresult-800ac472-from-set-operations-in-excel?forum=vsto
  constexpr HRESULT VBA_E_IGNORE = 0x800ac472;

  template <class TFunc>
  bool retryComCall(TFunc fn)
  {
    XLO_TRACE("Calling into XLL context fn= {0:#x}", (size_t)&fn);
    for (auto tries = 0; tries < 10; ++tries)
    {
      try
      {
        fn();
        return true;
      }
      catch (_com_error& error)
      {
        if (error.Error() != VBA_E_IGNORE)
        {
          XLO_ERROR(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
          break;
        }
      }
      Sleep(50);
      XLO_TRACE("Retry # {0} for COM call", (tries + 1));
    }
    return false;
  }
}