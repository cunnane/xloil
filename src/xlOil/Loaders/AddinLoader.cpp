#include "AddinLoader.h"
#include <xlOil/Interface.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Date.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Log.h>
#include <xlOil/Events.h>
#include <xloil/State.h>
#include <xloil/RtdServer.h>
#include <xloil-XLL/LogWindowSink.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelThread.h>
#include <xlOil-COM/Connect.h>
#define TOML_ABI_NAMESPACES 0
#include <toml++/toml.h>
#include <filesystem>

namespace fs = std::filesystem;
using std::make_pair;
using std::wstring;
using std::make_shared;
using std::shared_ptr;

namespace xloil
{
  namespace
  {
    std::map<std::wstring, std::shared_ptr<AddinContext>> theAddinContexts;
    std::shared_ptr<AddinContext> theCoreContextPtr;

    /// <summary>
    /// Finds the settings file <XllName>.ini either in %APPDATA%\xlOil
    /// or in the same directory as the XLL.  Adds any log sink specified
    /// and any date formats.
    /// </summary>
    auto createContext(const wchar_t* xllPath)
    {
      auto settings = findSettingsFile(xllPath);
      wstring logFile;
      if (!settings)
      {
        XLO_DEBUG(L"No settings file found for '{}'", xllPath);
      }
      else
      {
        // Log file settings
        logFile = Settings::logFilePath(*settings);
        auto logLevel = Settings::logLevel(*settings);
        auto [logMaxSize, logNumFiles] = Settings::logRotation(*settings);

        logFile = loggerAddRotatingFileSink(spdlog::default_logger(),
                                            logFile.c_str(), logLevel.c_str(),
                                            logMaxSize, logNumFiles);

        loggerSetFlush(spdlog::default_logger(),
                       Settings::logFlushLevel(*settings));

        // Write the log message *after* we set up the log file!
        XLO_INFO(L"Found core settings file '{}' for '{}'",
          utf8ToUtf16(*settings->source().path), xllPath);

        // If this is specified in multiple addins and/or the core, 
        // the last value overrides: not easy to workaround
        setLogWindowPopupLevel(
            Settings::logPopupLevel(*settings).c_str());

        // Add any requested date formats
        auto dateFormats = Settings::dateFormats(*settings);
        for (auto& form : dateFormats)
          theDateTimeFormats().push_back(form);
      }

      auto [ctx, isNew] = theAddinContexts.insert_or_assign(
        wstring(xllPath), make_shared<AddinContext>(xllPath, settings));

      ctx->second->logFilePath = std::move(logFile);
      return ctx->second;
    }
  }

  AddinContext& theCoreContext()
  {
    assert(theCoreContextPtr);
    return *theCoreContextPtr;
  }

  const std::map<std::wstring, std::shared_ptr<AddinContext>>& currentAddinContexts()
  {
    return theAddinContexts;
  }

  std::shared_ptr<AddinContext> createCoreAddinContext()
  {
    if (!theCoreContextPtr)
      theCoreContextPtr = createContext(Environment::coreDllPath());
    return theCoreContextPtr;
  }

  std::shared_ptr<AddinContext>createAddinContext(const wchar_t* pathName)
  {
    // Compare the filename stem to our core dll name (which should end in 'dll')
    const auto lastSlash = wcsrchr(pathName, L'\\');
    const auto coreDll = Environment::coreDllName();
    const bool isCore = 0 == _wcsnicmp(
      coreDll,
      lastSlash ? lastSlash + 1 : pathName,
      wcslen(coreDll) - 3);

    XLO_DEBUG(L"Creating addin context for '{}'", pathName);
    if (isCore)
    {
      auto context = createCoreAddinContext();
      // Point xloil.xll at the context (as well as xloil.dll). In the case
      // of multiple xlOil xlls, they can be unloaded in any order, the 
      // check in addinCloseXll that only 1 remains assumes there is 
      // one context entry per XLL.
      theAddinContexts[pathName] = theCoreContextPtr;
      return context;
    }
    else
      return createContext(pathName);
  }

  void addinCloseXll(const wchar_t* xllPath)
  {
    theAddinContexts.erase(xllPath);
    // Check if only the core is left
    if (theAddinContexts.size() == 1)
    {
      // Somewhat cheap trick to ensure any async tasks which may reference plugin
      // code are destroyed in a timely manner prior to teardown.  Better would be
      // to keep track of which tasks were registered by which addin
      rtdAsyncServerClear();

      if (theAddinContexts.begin()->second.get() != theCoreContextPtr.get())
        XLO_ERROR("addinCloseXll: unexpected addins present");

      theAddinContexts.clear();
      theCoreContextPtr->detachPlugins();
      theCoreContextPtr.reset();

      Event::AutoClose().fire();

      unloadAllPlugins();

      // We don't want any messages hanging around after autoClose finishes
      teardownMessageQueue();

      COM::disconnectCom();
    }

    spdlog::default_logger()->flush();
  }

  void teardownAddinContext()
  {
    teardownFunctionRegistry();

    theAddinContexts.clear();
    theCoreContextPtr.reset();

    teardownMessageQueue();

    COM::disconnectCom();

    spdlog::default_logger()->flush();
  }
}