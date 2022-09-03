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
#include <xlOil-COM/Connect.h>
#include <tomlplusplus/toml.hpp>
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

    /// <summary>
    /// Finds the settings file <XllName>.ini either in %APPDATA%\xlOil
    /// or in the same directory as the XLL.  Adds any log sink specified
    /// and any date formats.
    /// </summary>
    auto processAddinSettings(const wchar_t* xllPath)
    {
      auto settings = findSettingsFile(xllPath);
      if (!settings)
      {
        XLO_DEBUG(L"No settings file found for '{}'", xllPath);
        return settings;
      }

      auto addinRoot = (*settings)[XLOIL_SETTINGS_ADDIN_SECTION];

      // Log file settings
      auto logFile = Settings::logFilePath(*settings);
      auto logLevel = Settings::logLevel(addinRoot);
      auto [logMaxSize, logNumFiles] = Settings::logRotation(addinRoot);

      loggerAddRotatingFileSink(
        spdlog::default_logger(),
        logFile.c_str(), logLevel.c_str(), 
        logMaxSize, logNumFiles);

      // Write the log message *after* we set up the log file!
      XLO_INFO(L"Found core settings file '{}' for '{}'",
        utf8ToUtf16(*settings->source().path), xllPath);

      // If this is specified in multiple addins and/or the core, 
      // the last value overrides: not easy to workaround
      setLogWindowPopupLevel(
        spdlog::level::from_str(
          Settings::logPopupLevel(addinRoot).c_str()));

      // Add any requested date formats
      auto dateFormats = Settings::dateFormats(addinRoot);
      for (auto& form : dateFormats)
        theDateTimeFormats().push_back(form);

      return settings;
    }
  }

  static xloil::AddinContext* ourCoreContext;

  AddinContext& theCoreContext()
  {
    return *ourCoreContext;
  }

  const std::map<std::wstring, std::shared_ptr<AddinContext>>& currentAddinContexts()
  {
    return theAddinContexts;
  }

  void createCoreContext() 
  {
    ourCoreContext = &createAddinContext(Environment::coreDllPath());
    auto staticSource = make_shared<StaticFunctionSource>(Environment::coreDllName());
    ourCoreContext->addSource(staticSource);
  }

  AddinContext& createAddinContext(const wchar_t* pathName)
  {
    auto settings = processAddinSettings(pathName);
    auto [ctx, isNew] = theAddinContexts.insert_or_assign(
      wstring(pathName), make_shared<AddinContext>(pathName, settings));

    return *ctx->second;
  }

  void addinCloseXll(const wchar_t* xllPath)
  {
    theAddinContexts.erase(xllPath);
    // Check if only the core left
    if (theAddinContexts.size() == 1)
    {
      theAddinContexts.clear();
      
      // Somewhat cheap trick to ensure any async tasks which may reference plugin
      // code are destroyed in a timely manner prior to teardown.  Better would be
      // to keep track of which tasks were registered by which addin
      rtdAsyncServerClear();

      Event::AutoClose().fire();

      unloadAllPlugins();
      assert(theAddinContexts.empty());

      COM::disconnectCom();
    }
  }
}