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
#include <filesystem>

namespace fs = std::filesystem;

using std::wstring;
using std::string;
using std::vector;
using std::shared_ptr;
using std::tuple;

namespace
{
  static HMODULE theCoreModuleHandle = nullptr;
  static bool theCoreIsLoaded = false;
}

namespace xloil
{
  XLOIL_EXPORT int coreAutoOpenHandler(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;
      // A return val of 1 tells the XLL to hook XLL-API events. There may be
      // multiple XLLs calling this function, but we only want to hook the events 
      // the first time.
      int retVal = 0;

      shared_ptr<FuncSource> coreRegisteredFunctions;
      std::shared_ptr<spdlog::logger> logger;

      if (!theCoreIsLoaded)
      {
        Environment::initAppContext();
        Environment::setCoreHandle(theCoreModuleHandle);

        // There's no log file until createAddinContext figures out our 
        // settings, so any logging goes to the debug output.  We also flush
        // on trace level so we don't miss any crashes during startup. This
        // has a minimal performance impact vs flushing during sheet calc.
        logger = loggerInitialise("trace");
        XLO_INFO(L"xlOil {} starting", XLOIL_VERSION_STR);
        loggerSetFlush(logger, "trace");

        initMessageQueue(Environment::excelProcess().hInstance);

        loggerAddPopupWindowSink(logger);
      }

      // After the context has been created, we will have a log file
      auto coreContext = createCoreAddinContext();
      auto addinContext = createAddinContext(xllPath);

      if (!theCoreIsLoaded)
      {
        XLO_DEBUG(L"Loaded xlOil core from: {}", Environment::coreDllPath());

        // Flush logger after sheet calculates
        Event::AfterCalculate() += [logger]() { logger->flush(); };

        // Run before staticSource so the function registration gets picked up
        registerIntellisenseHook(xllPath);

        // Collect all static UDFs for registration
        coreRegisteredFunctions = std::make_shared<StaticFunctionSource>(
          Environment::coreDllName());

        // Do the registration
        coreRegisteredFunctions->init();

        // Associate registed functions with the core 
        coreContext->addSource(coreRegisteredFunctions);
        // Signal that the XLL events should be hooked
        retVal = 1;
      }

      if (addinContext == coreContext || theCoreIsLoaded)
      {
        addinContext->loadPlugins();
      }
      else
      {
        // Check if we should process the settings for a non-core addin first
        // and/or we need to load the core addin. We also check we don't call
        // loadPluginsForAddin twice (although it would be harmless)
        const bool loadBeforeCore = Settings::loadBeforeCore(*addinContext->settings());

        const auto [firstLoad, secondLoad] = loadBeforeCore
          ? tuple(addinContext, coreContext)
          : tuple(coreContext, addinContext);

        XLO_DEBUG(L"User xll file present, plugin load order is: {}, {}",
          firstLoad ? firstLoad->pathName() : L"", 
          firstLoad ? secondLoad->pathName() : L"");

        if (firstLoad)
          firstLoad->loadPlugins();
        if (secondLoad)
          secondLoad->loadPlugins();
      }

      runComSetupOnXllOpen([]() {});

      theCoreIsLoaded = true;
      return retVal;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Initialisation error: {0}", e.what());
    }

    return 0;
  }
  XLOIL_EXPORT int coreAutoCloseHandler(const wchar_t* xllPath) noexcept
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

XLO_DEFINE_FREE_CALLBACK()

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
  else if (fdwReason == DLL_PROCESS_DETACH)
  {
    xloil::teardownAddinContext();
    theCoreModuleHandle = nullptr;
  }
  return TRUE;
}
