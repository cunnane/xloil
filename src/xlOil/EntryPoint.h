#pragma once
#include "ExportMacro.h"
#include "WindowsSlim.h"
#include <delayimp.h>

namespace xloil
{
  class Core;
  typedef FARPROC(*coreLoadHook)(unsigned dliNotify, PDelayLoadInfo pdli);

  void* coreModuleHandle();

  const wchar_t* theCorePath();
  
  const wchar_t* theCoreName();

  const wchar_t* theXllPath();

  XLOIL_EXPORT int coreInit(coreLoadHook coreLoaderHook, const wchar_t* xllPath) noexcept;
  XLOIL_EXPORT int coreExit() noexcept;
}