#pragma once
#include <xloil/ExportMacro.h>

namespace xloil
{
  /// <summary>
  /// Called by the XLL loader's xlAutoOpen
  /// </summary>
  XLOIL_EXPORT int 
    autoOpenHandler(const wchar_t* xllPath) noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoClose
  /// </summary>
  XLOIL_EXPORT int 
    autoCloseHandler(const wchar_t* xllPath) noexcept;

  XLOIL_EXPORT void onCalculationCancelled() noexcept;
}