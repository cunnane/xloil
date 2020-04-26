#include <xlOil/Interface.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/EntryPoint.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Log.h>
#include <xlOil/Events.h>
#include <toml11/toml.hpp>
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
        const wchar_t* pathName, std::shared_ptr<const toml::value> settings)
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
    return make_pair(std::shared_ptr<FileSource>(), std::shared_ptr<AddinContext>());
  }

  void
    deleteFileSource(const std::shared_ptr<FileSource>& context)
  {
    for (auto[name, addinCtx] : theAddinContexts)
    {
      auto found = addinCtx->files().find(context->sourceName());
      if (found != addinCtx->files().end())
        return addinCtx->removeFileSource(found);
    }
  }

  void openXll(const wchar_t* xllPath)
  {
    // On First load, register the core functions
    if (theAddinContexts.empty())
    {
      auto settings = findSettingsFile(theCorePath());
      if (settings)
        XLO_DEBUG("Found core settings file '{0}'", 
          settings->location().file_name());

      ourCoreContext = createAddinContext(theCorePath(), settings);

      auto logFile = Settings::logFilePath(settings.get());
      loggerInitialise(Settings::logLevel(settings.get()).c_str());
      loggerAddFile(logFile.empty()
        ? fs::path(theCorePath()).replace_extension("log").c_str()
        : logFile.c_str());
      
      ourCoreContext->tryAdd<StaticFunctionSource>(theCoreName(), theCoreName());

      loadPlugins(ourCoreContext, Settings::plugins(settings.get()));
    }

    if (_wcsicmp(fs::path(xllPath).replace_extension("dll").c_str(),
      theCorePath()) == 0)
      return;

    auto settings = findSettingsFile(xllPath);
    if (settings)
      XLO_DEBUG("Found settings file '{0}'", 
        settings->location().file_name());

    
    if (theAddinContexts.find(xllPath) != theAddinContexts.end())
      theAddinContexts.erase(xllPath);
    
    auto ctx = createAddinContext(xllPath, settings);
    assert(ctx);

    auto logFile = Settings::logFilePath(settings.get());
    loggerAddFile(logFile.empty()
      ? fs::path(xllPath).replace_extension("log").c_str()
      : logFile.c_str());

    loadPlugins(ctx, Settings::plugins(settings.get()));
  }

  void closeXll(const wchar_t* xllPath)
  {
    theAddinContexts.erase(xllPath);

    // Check if only the core left
    if (theAddinContexts.size() == 1)
    {
      theAddinContexts.erase(theCorePath());
      // TODO: remove this legacy event?
      Event_AutoClose().fire();
      unloadPlugins();
      assert(theAddinContexts.empty());
    }
  }
}