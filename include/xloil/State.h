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

    struct ExcelState
    {
      /// <summary>
      /// The Excel major version number
      /// </summary>
      int version;
      /// <summary>
      /// The Windows API process instance handle, castable to HINSTANCE
      /// </summary>
      void* hInstance;
      /// <summary>
      /// The Windows API handle for the top level Excel window 
      /// castable to type HWND
      /// </summary>
      long long hWnd;
      size_t mainThreadId;
    };

    /// <summary>
    /// Returns Excel application state information such as the version number,
    /// HINSTANCE, window handle and thread ID.
    /// </summary>
    XLOIL_EXPORT ExcelState& excelState() noexcept;

    void initCoreContext(void* coreHInstance);
    XLOIL_EXPORT void initAppContext();
  }
}