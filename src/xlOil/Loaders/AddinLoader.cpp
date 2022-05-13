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
#include <xloil/LogWindowSink.h>
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

    auto processAddinSettings(const wchar_t* xllPath)
    {
      auto settings = findSettingsFile(xllPath);
      if (!settings)
      {
        XLO_DEBUG(L"No settings file found for '{}'", xllPath);
        return settings;
      }

      auto addinRoot = (*settings)["Addin"];

      // Log file settings
      auto logFile = Settings::logFilePath(*settings);
      auto logLevel = Settings::logLevel(addinRoot);
      auto [logMaxSize, logNumFiles] = Settings::logRotation(addinRoot);

      detail::loggerAddFile(
        logFile.c_str(), logLevel.c_str(), 
        logMaxSize, logNumFiles);

      XLO_INFO("Found core settings file '{}'",
        *settings->source().path);

      // Add any requested date formats
      auto dateFormats = Settings::dateFormats(addinRoot);
      for (auto& form : dateFormats)
        dateTimeAddFormat(form.c_str());

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
    ourCoreContext = &createAddinContext(State::coreDllPath());

    const auto& coreAddinSettings = (*ourCoreContext->settings())["Addin"];

    // Can only do this once not per-addin
    setLogWindowPopupLevel(
      spdlog::level::from_str(
        Settings::logPopupLevel(coreAddinSettings).c_str()));

    auto staticSource = make_shared<StaticFunctionSource>(State::coreDllName());
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