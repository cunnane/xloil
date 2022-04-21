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

    AddinContext* createAddinContext(
      const wchar_t* pathName, const std::shared_ptr<const toml::table>& settings)
    {
      auto [ctx, isNew] = theAddinContexts.try_emplace(
        wstring(pathName), make_shared<AddinContext>(pathName, settings));
      return isNew ? ctx->second.get() : nullptr;
    }

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

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    findFileSource(const wchar_t* source)
  {
    for (auto& [addinName, addin] : theAddinContexts)
    {
      auto found = addin->files().find(source);
      if (found != addin->files().end())
        return make_pair(found->second, addin);
    }
    return make_pair(shared_ptr<FileSource>(), shared_ptr<AddinContext>());
  }

  void
    deleteFileSource(const std::shared_ptr<FileSource>& context)
  {
    for (auto& [name, addinCtx] : theAddinContexts)
    {
      auto found = addinCtx->files().find(context->sourcePath());
      if (found != addinCtx->files().end())
        addinCtx->removeSource(found);
    }
  }

  namespace
  {
    AddinContext& createAddinContext(const wchar_t* addinPathName)
    {
      // Delete existing context if addin is reloaded
      if (theAddinContexts.find(addinPathName) != theAddinContexts.end())
        theAddinContexts.erase(addinPathName);

      auto settings = processAddinSettings(addinPathName);
      auto ctx = createAddinContext(addinPathName, settings);
      if (!ctx)
        XLO_THROW(L"Failed to create add-in context for {}", addinPathName);

      return *ctx;
    }
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
    staticSource->registerQueue();
    ourCoreContext->addSource(staticSource);
  }

  void loadPluginsForAddin(AddinContext& ctx)
  {
    auto plugins = Settings::plugins((*ctx.settings())["Addin"]);
    loadPlugins(ctx, plugins);
  }

  AddinContext& addinOpenXll(const wchar_t* xllPath)
  {
    auto& ctx = createAddinContext(xllPath);
    return ctx;
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