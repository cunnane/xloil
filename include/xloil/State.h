#pragma once
#include <xlOil/ExportMacro.h>

namespace xloil
{
  namespace Environment
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

    struct ExcelProcessInfo
    {
      ExcelProcessInfo();

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
      /// <summary>
      /// Thread Id of Excel's main thread
      /// </summary>
      size_t mainThreadId;
    };

    /// <summary>
    /// Returns Excel application state information such as the version number,
    /// HINSTANCE, window handle and thread ID.
    /// </summary>
    XLOIL_EXPORT const ExcelProcessInfo& excelProcess() noexcept;

    /// <summary>
    /// Internal usage
    /// </summary>
    void initCoreContext(void* coreHInstance);
    /// <summary>
    /// Internal usage
    /// </summary>
    XLOIL_EXPORT void initAppContext();
  }
}