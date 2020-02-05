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
#include "COMInterface/Connect.h"
#include <boost/preprocessor/cat.hpp>
#include <boost/preprocessor/repeat.hpp>
#include <delayimp.h>

using std::wstring;

namespace
{
  xloil::coreLoadHook theCoreLoader = nullptr;

  FARPROC coreLoadThunk(unsigned dliNotify, PDelayLoadInfo pdli)
  {
    return theCoreLoader(dliNotify, pdli);
  }
}

extern "C" const PfnDliHook __pfnDliNotifyHook2 = coreLoadThunk;

namespace
{
  static HMODULE theCoreModuleHandle = nullptr;

  static const wchar_t* ourDllName;
  static wchar_t ourDllPath[4 * MAX_PATH]; // TODO: may not be long enough!!!
  static const wchar_t* ourXllPath = nullptr;

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

  XLOIL_EXPORT int coreInit(
    coreLoadHook coreLoaderHook, 
    const wchar_t* xllPath) noexcept
  {
    try
    {
      ScopeInXllContext xllContext;

      theCoreLoader = coreLoaderHook;
      //__HrLoadAllImportsForDll("Core.dll");
      ourXllPath = xllPath;
      auto& settings = theCoreSettings();

      initialiseLogger(settings.logLevel, settings.logFilePath.empty() ? nullptr : &settings.logFilePath);
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
  XLOIL_EXPORT int coreExit() noexcept
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

#define XLO_WRITE_STUB(z, n, dummy) extern "C"  __declspec(dllexport) void* __stdcall XLOIL_STUB(n)() { return nullptr; }
BOOST_PP_REPEAT(XLOIL_MAX_FUNCS, XLO_WRITE_STUB, 0)
#undef XLO_WRITE_STUB