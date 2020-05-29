#pragma once
#include <xloil/ExportMacro.h>
namespace Excel { struct _Application; }

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
    XLOIL_EXPORT const wchar_t* corePath() noexcept;

    /// <summary>
    /// Name of the xlOil Core DLL including the extension 
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreName() noexcept;

    /// <summary>
    /// Returns the Excel major version number
    /// </summary>
    XLOIL_EXPORT int excelVersion() noexcept;

    XLOIL_EXPORT void* excelHInstance() noexcept;

    XLOIL_EXPORT Excel::_Application& excelApp() noexcept;

    void initAppContext(void* coreHInstance);
  }
}