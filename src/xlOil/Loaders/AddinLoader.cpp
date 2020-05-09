#include <xlOil/Interface.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Date.h>
#include <xlOil/EntryPoint.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Log.h>
#include <xlOil/Events.h>
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

  auto processAddinSettings(const wchar_t* xllPath)
  {
    auto settings = findSettingsFile(theCorePath());
    if (!settings)
      return settings;

    XLO_DEBUG("Found core settings file '{0}'",
      *settings->source().path);

    // Log file settings
    auto logFile = Settings::logFilePath(settings.get());
    auto logLevel = Settings::logLevel(settings.get());
    if (logFile.empty())
      logFile = fs::path(xllPath).replace_extension("log");
    loggerAddFile(logFile.c_str(), logLevel.c_str());

    // Add any requested date formats
    auto dateFormats = Settings::dateFormats(settings.get());
    for (auto& form : dateFormats)
      dateTimeAddFormat(form.c_str());

    return settings;
  }

  void openXll(const wchar_t* xllPath)
  {
    // On First load, register the core functions
    if (theAddinContexts.empty())
    {
#if _DEBUG
      loggerInitialise(spdlog::level::debug);
#else
      loggerInitialise(spdlog::level::warn);
#endif

      auto settings = processAddinSettings(theCorePath());
      
      ourCoreContext = createAddinContext(theCorePath(), settings);
      ourCoreContext->tryAdd<StaticFunctionSource>(theCoreName(), theCoreName());

      loadPlugins(ourCoreContext, Settings::plugins(settings.get()));
    }

    if (_wcsicmp(fs::path(xllPath).replace_extension("dll").c_str(),
      theCorePath()) == 0)
      return;

    auto settings = processAddinSettings(xllPath);
    
    // Delete existing context if addin is reloaded
    if (theAddinContexts.find(xllPath) != theAddinContexts.end())
      theAddinContexts.erase(xllPath);
    
    auto ctx = createAddinContext(xllPath, settings);
    assert(ctx);

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
      Event::AutoClose().fire();
      unloadPlugins();
      assert(theAddinContexts.empty());
    }
  }
}