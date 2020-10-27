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
    XLOIL_EXPORT const wchar_t* corePath() noexcept;

    /// <summary>
    /// Name of the xlOil Core DLL including the extension 
    /// </summary>
    XLOIL_EXPORT const wchar_t* coreName() noexcept;

    struct ExcelState
    {
      int version;
      void* hInstance;
      long long hWnd;
      size_t mainThreadId;
    };

    /// <summary>
    /// Returns the Excel major version number
    /// </summary>
    XLOIL_EXPORT ExcelState& excelState() noexcept;

    void initAppContext(void* coreHInstance);
  }
}