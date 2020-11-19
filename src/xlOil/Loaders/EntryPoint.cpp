#include <xlOil/Events.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/Interface.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/Log.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/WindowsSlim.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xlOil/State.h>
#include <xlOil/ExcelApp.h>
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
      // mulltiple XLLs, but we only want to hook the events once, when we load 
      // the core DLL.
      int retVal = 0;

      if (!theCoreIsLoaded)
      {
#if _DEBUG
        detail::loggerInitialise(spdlog::level::debug);
#else
        detail::loggerInitialise(spdlog::level::err);
#endif
        State::initCoreContext(theCoreModuleHandle);

        detail::loggerInitPopupWindow();

        createCoreContext();

        xllOpenComCall([&]() { loadPluginsForAddin(theCoreContext()); });

        theCoreIsLoaded = true;
        retVal = 1;

        // TODO: no need to call this
        initMessageQueue();
      }

      if (_wcsicmp(L"xloil.xll", fs::path(xllPath).filename().c_str()) != 0)
      {
        auto addinContext = openXll(xllPath);
        xllOpenComCall([=]() { loadPluginsForAddin(addinContext); });
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
      
      closeXll(xllPath);

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

