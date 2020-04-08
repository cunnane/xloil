#include "Options.h"
#include "Events.h"
#include "Loader.h"
#include "WindowsSlim.h"
#include "ExcelObj.h"
#include "Interface.h"
#include "ExcelCall.h"
#include "EntryPoint.h"
#include "ExportMacro.h"
#include "Log.h"
#include "Settings.h"
#include <COMInterface/Connect.h>
#include <COMInterface/XllContextInvoke.h>
#include <delayimp.h>

using std::wstring;

namespace
{
  static HMODULE theCoreModuleHandle = nullptr;

  static const wchar_t* ourDllName;
  static wchar_t ourDllPath[4 * MAX_PATH]; // TODO: may not be long enough!!!
  static const wchar_t* ourXllPath = nullptr;
  static int ourExcelVersion;

  bool setDllPath(HMODULE handle)
  {
    auto size = GetModuleFileName(handle, ourDllPath, sizeof(ourDllPath));
    if (size == 0)
    {
      XLO_ERROR(L"Could not determine XLL location: {}", xloil::writeWindowsError());
      return false;
    }
    ourDllName = wcsrchr(ourDllPath, L'\\') + 1;
    return true;
  }
}

namespace xloil
{
  void* coreModuleHandle()
  {
    return theCoreModuleHandle;
  }
  const wchar_t* theCorePath()
  {
    return ourDllPath;
  }
  const wchar_t* theCoreName()
  {
    return ourDllName;
  }
  const wchar_t* theXllPath()
  {
    return ourXllPath;
  }
  int coreExcelVersion()
  {
    return ourExcelVersion;
  }

  int getExcelVersion()
  {
    // https://github.com/MicrosoftDocs/office-developer-client-docs/blob/...
    // master/docs/excel/calling-into-excel-from-the-dll-or-xll.md
    auto versionStr = callExcel(msxll::xlfGetWorkspace, 2);
    return std::stoi(versionStr.toString());
  }

  XLOIL_EXPORT int coreInit(const wchar_t* xllPath) noexcept
  {
    ourXllPath = xllPath;
    return 1;
  }

  XLOIL_EXPORT int coreAutoOpen() noexcept
  {
    try
    {
      ScopeInXllContext xllContext;
      

      auto& settings = theCoreSettings();

      initialiseLogger(settings.logLevel, settings.logFilePath.empty() 
        ? nullptr : &settings.logFilePath);
      
      ourExcelVersion = getExcelVersion();
      loadPlugins();
      excelApp(); // Creates the COM connection

      return 1;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Initialisation error: {0}", e.what());
    }
    return 0;
  }
  XLOIL_EXPORT int coreAutoClose() noexcept
  {
    try
    {
      ScopeInXllContext xllContext;
      unloadPlugins();
      Event_AutoClose().fire();
      return 1;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Finalisation error: {0}", e.what());
    }
    return 0;
  }
}

XLO_ENTRY_POINT(int) DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    /*lpvReserved*/
)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
  {
    theCoreModuleHandle = hinstDLL;
    if (!setDllPath(hinstDLL))
      return FALSE;
  }
  return TRUE;
}

extern "C"  __declspec(dllexport) void* __stdcall XLOIL_STUB_NAME() 
{ 
  return nullptr; 
}
