#pragma once
#include "ExportMacro.h"

#ifdef _WIN64
typedef __int64 (_stdcall *FARPROC)();
#else
typedef int ( _stdcall *FARPROC)();
#endif  // _WIN64

struct DelayLoadInfo;

namespace xloil
{
  class Core;
  typedef FARPROC(*coreLoadHook)(unsigned dliNotify, DelayLoadInfo* pdli);
  
  void* coreModuleHandle();

  const wchar_t* theCorePath();
  
  const wchar_t* theCoreName();

  int coreExcelVersion();

  /// <summary>
  /// Path to the xll loaded by Excel, not the core DLL
  /// </summary>
  const wchar_t* theXllPath();
  XLOIL_EXPORT int coreInit(const wchar_t* xllPath) noexcept;
  XLOIL_EXPORT int coreAutoOpen() noexcept;
  XLOIL_EXPORT int coreAutoClose() noexcept;
}