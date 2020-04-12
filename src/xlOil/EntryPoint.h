#pragma once
#include "ExportMacro.h"

namespace xloil
{
  /// <summary>
  /// The HINSTANCE for this DLL, as passed into DllMain
  /// </summary>
  void* coreModuleHandle();

  /// <summary>
  /// Path to the xlOil Core DLL, including the DLL name
  /// </summary>
  const wchar_t* theCorePath();
  
  /// <summary>
  /// Name of the xlOil Core DLL including the extension 
  /// </summary>
  const wchar_t* theCoreName();

  /// <summary>
  /// Returns the Excel major version number
  /// </summary>
  int coreExcelVersion();

  /// <summary>
  /// Path to the xll loaded by Excel, not the core DLL
  /// </summary>
  const wchar_t* theXllPath();

  /// <summary>
  /// Called by the XLL loader's DllMain, passing the path
  /// to the loader.
  /// </summary>
  XLOIL_EXPORT int 
    coreInit(const wchar_t* xllPath) noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoOpen
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoOpen() noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoClose
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoClose() noexcept;
}