#pragma once
#include <xlOil/ExportMacro.h>

namespace xloil
{
   namespace State
  {
    /// <summary>
    /// The HINSTANCE for this DLL, as passed into DllMain
    /// </summary>
    XLOIL_EXPORT void* coreModuleHandle() noexcept;

    /// <summary>
    /// Path to the xlOil Core DLL, including the DLL name
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreDllPath() noexcept;

    /// <summary>
    /// Name of the xlOil Core DLL including the extension 
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreDllName() noexcept;

    void initCoreContext(void* coreHInstance);
    XLOIL_EXPORT void initAppContext();
  }
}