#include <xlOil/Interface.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOil/Loaders/Settings.h>
#include <xlOil/EntryPoint.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOil/Log.h>
#include <xlOil/Events.h>
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
      auto ctx = shared_ptr<AddinContext>(new AddinContext(pathName, settings));
      theAddinContexts.emplace(make_pair(wstring(pathName), ctx));
      return ctx.get();
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
      auto logFile = Settings::logFilePath(settings.get());
      initialiseLogger(Settings::logLevel(settings.get()), logFile.empty()
        ? nullptr : &logFile);
      ourCoreContext = createAddinContext(theCorePath(), settings);
      ourCoreContext->tryAdd<StaticFunctionSource>(theCoreName(), theCoreName());

      loadPlugins(ourCoreContext, Settings::plugins(settings.get()));
    }

    if (_wcsicmp(fs::path(xllPath).filename().c_str(), L"xlOil.xll") != 0)
    {
      auto settings = findSettingsFile(xllPath);
      auto ctx = createAddinContext(xllPath, settings);
      loadPlugins(ctx, Settings::plugins(settings.get()));
    }
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