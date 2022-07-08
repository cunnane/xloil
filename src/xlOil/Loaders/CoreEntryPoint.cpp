#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/CoreEntryPoint.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xlOil/State.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/Interface.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/Log.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil-XLL/Intellisense.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/XllContextInvoke.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::wstring;
using std::string;
using std::vector;
using std::shared_ptr;

namespace
{
  static HMODULE theCoreModuleHandle = nullptr;
  static bool theCoreIsLoaded = false;
}

namespace xloil
{
  XLOIL_EXPORT int autoOpenHandler(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;
      // A return val of 1 tells the XLL to hook XLL-api events. There may be
      // multiple XLLs, but we only want to hook the events once, when we load 
      // the core DLL.
      int retVal = 0;

      if (!theCoreIsLoaded)
      {
        // There's no log file until createAddinContext figures out our 
        // settings, so any logging goes to the debug output.
        detail::loggerInitialise(spdlog::level::debug);

        Environment::initCoreContext(theCoreModuleHandle);

        XLO_DEBUG(L"Loaded xlOil core from: {}", Environment::coreDllPath());

        detail::loggerInitPopupWindow();
      }

      bool isXloilCoreAddin = _wcsicmp(L"xloil.xll", fs::path(xllPath).filename().c_str()) == 0;
      AddinContext* addinContext = nullptr;

      if (!isXloilCoreAddin)
      {
        addinContext = &createAddinContext(xllPath);
        auto loadFirst = Settings::loadBeforeCore(*addinContext->settings());
        if (loadFirst)
          runComSetupOnXllOpen([&]() { loadPluginsForAddin(*addinContext); });
        addinContext = nullptr;
      }

      if (!theCoreIsLoaded)
      {
        // Run *before* createCoreContext so the function registration memo gets
        // picked up
        registerIntellisenseHook(xllPath);

        createCoreContext();

        runComSetupOnXllOpen([&]() { loadPluginsForAddin(theCoreContext()); });

        theCoreIsLoaded = true;
        retVal = 1;
      }

      // If we have an addin context here it means we are not
      // xloil.xll and have not yet loaded our plugins
      if (addinContext)
      {
        runComSetupOnXllOpen([&]() { loadPluginsForAddin(*addinContext); });
      }

      return retVal;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Initialisation error: {0}", e.what());
    }
    return -1;
  }
  XLOIL_EXPORT int autoCloseHandler(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;
      
      addinCloseXll(xllPath);

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
  }
  return TRUE;
}

