#pragma once
#include <xloil/ExportMacro.h>

#define XLOIL_STUB_NAME xloil_stub

namespace xloil
{
  /// <summary>
  /// Called by the XLL loader's xlAutoOpen
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoOpen(const wchar_t* xllPath) noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoClose
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoClose(const wchar_t* xllPath) noexcept;

  XLOIL_EXPORT void onCalculationEnded() noexcept;
  XLOIL_EXPORT void onCalculationCancelled() noexcept;
}