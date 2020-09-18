#include "AddinLoader.h"
#include <xlOil/Interface.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Date.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Log.h>
#include <xlOil/Events.h>
#include <xloil/State.h>
#include <xloil/RtdServer.h>
#include <xloil/LogWindowSink.h>
#include <COMInterface/Connect.h>
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

    AddinContext*
      createAddinContext(
        const wchar_t* pathName, std::shared_ptr<const toml::table> settings)
    {
      auto [ctx, isNew] = theAddinContexts.try_emplace(
        wstring(pathName), make_shared<AddinContext>(pathName, settings));
      return isNew ? ctx->second.get() : nullptr;
    }
  }

  static xloil::AddinContext* ourCoreContext;

  AddinContext* theCoreContext()
  {
    return ourCoreContext;
  }

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    findFileSource(const wchar_t* source)
  {
    for (auto&[addinName, addin] : theAddinContexts)
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
    for (auto[name, addinCtx] : theAddinContexts)
    {
      auto found = addinCtx->files().find(context->sourcePath());
      if (found != addinCtx->files().end())
        addinCtx->removeSource(found);
    }
  }

  auto processAddinSettings(const wchar_t* xllPath)
  {
    auto settings = findSettingsFile(xllPath);
    if (!settings)
      return settings;

    XLO_DEBUG("Found core settings file '{0}'",
      *settings->source().path);

    auto addinRoot = (*settings)["Addin"];

    // Log file settings
    auto logFile = Settings::logFilePath(*settings);
    auto logLevel = Settings::logLevel(addinRoot);
    detail::loggerAddFile(logFile.c_str(), logLevel.c_str());

    // Add any requested date formats
    auto dateFormats = Settings::dateFormats(addinRoot);
    for (auto& form : dateFormats)
      dateTimeAddFormat(form.c_str());

    return settings;
  }

  void createCoreContext() 
  {
    ourCoreContext = openXll(State::corePath());
    // Can only do this once not per-addin
    setLogWindowPopupLevel(
      spdlog::level::from_str(
        Settings::logPopupLevel((*ourCoreContext->settings())["Addin"]).c_str()));

    ourCoreContext->tryAdd<StaticFunctionSource>(State::coreName(), State::coreName());
  }

  void loadPluginsForAddin(AddinContext* ctx)
  {
    auto plugins = Settings::plugins((*ctx->settings())["Addin"]);
    loadPlugins(ctx, plugins);
  }

  AddinContext* openXll(const wchar_t* xllPath)
  {
    auto settings = processAddinSettings(xllPath);
    
    // Delete existing context if addin is reloaded
    if (theAddinContexts.find(xllPath) != theAddinContexts.end())
      theAddinContexts.erase(xllPath);
    
    auto ctx = createAddinContext(xllPath, settings);
    assert(ctx);
    return ctx;
  }

  void closeXll(const wchar_t* xllPath)
  {
    theAddinContexts.erase(xllPath);

    // Check if only the core left
    if (theAddinContexts.size() == 1)
    {
      theAddinContexts.erase(State::corePath());
      
      // Somewhat cheap trick to ensure any async tasks which may reference plugin
      // code are destroyed in a timely manner prior to teardown.  Better would be
      // to keep track of which tasks were registered by which addin
      rtdAsyncServerClear();

      // TODO: remove this event?
      Event::AutoClose().fire();

      unloadAllPlugins();
      assert(theAddinContexts.empty());

      COM::disconnectCom();
    }
  }
}