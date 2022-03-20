#pragma once
#include <xlOil/ExportMacro.h>

namespace xloil
{
   namespace State
  {
    /// <summary>
    /// The HINSTANCE for the xlOil Core DLL, as passed into DllMain
    /// If xlOil is statically linked, the core is the DLL or XLL which 
    /// linked it.
    /// </summary>
    XLOIL_EXPORT void* coreModuleHandle() noexcept;

    /// <summary>
    /// Path to the xlOil Core DLL, including the DLL name.
    /// If xlOil is statically linked, the core is the DLL or XLL which 
    /// linked it.
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreDllPath() noexcept;

    /// <summary>
    /// Name of the xlOil Core DLL including the extension.
    /// If xlOil is statically linked, the core is the DLL or XLL which 
    /// linked it.
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreDllName() noexcept;

    void initCoreContext(void* coreHInstance);
    XLOIL_EXPORT void initAppContext();
  }
}