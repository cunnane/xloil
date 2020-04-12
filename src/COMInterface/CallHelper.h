#pragma once
#include "xloil/Log.h"
#include <optional>

namespace xloil
{
  // See https://social.msdn.microsoft.com/Forums/vstudio/en-US/9168f9f2-e5bc-4535-8d7d-4e374ab8ff09/hresult-800ac472-from-set-operations-in-excel?forum=vsto
  constexpr HRESULT VBA_E_IGNORE = 0x800ac472;

  template <class TFunc>
  auto retryComCall(TFunc fn, size_t nTries = 5) -> std::optional<typename std::invoke_result<TFunc>::type>
  {
    for (size_t tries = 0; tries < nTries; ++tries)
    {
      try
      {
        return std::optional<std::invoke_result<TFunc>::type>(fn());
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
      XLO_DEBUG("Retry # {0} for COM call", (tries + 1));
    }

    XLO_ERROR(L"Excel COM is busy. A dialog box may be open. If this error persists, restart Excel.");
    reconnectCOM();
    return std::nullopt;
  }
}