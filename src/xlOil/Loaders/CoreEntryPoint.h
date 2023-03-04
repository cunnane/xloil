#pragma once
#include <xloil/ExportMacro.h>

namespace xloil
{
  /// <summary>
  /// Called by the XLL loader's xlAutoOpen
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoOpenHandler(const wchar_t* xllPath) noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoClose
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoCloseHandler(const wchar_t* xllPath) noexcept;

  XLOIL_EXPORT void onCalculationCancelled() noexcept;
}