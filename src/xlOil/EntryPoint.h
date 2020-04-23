#pragma once
#include "ExportMacro.h"

#define XLOIL_STUB_NAME xloil_stub

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
  /// Called by the XLL loader's xlAutoOpen
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoOpen(const wchar_t* xllPath) noexcept;

  /// <summary>
  /// Called by the XLL loader's xlAutoClose
  /// </summary>
  XLOIL_EXPORT int 
    coreAutoClose(const wchar_t* xllPath) noexcept;

}