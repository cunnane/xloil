#include "PluginLoader.h"
#include <xlOil/WindowsSlim.h>
#include <xlOilHelpers/Environment.h>
#include <xlOil-COM/RtdManager.h>
#include <xlOil/Log.h>
#include <xlOil/Throw.h>
#include <xlOil/State.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/StaticRegister.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xlOil/Version.h>
#include <toml++/toml.h>
#include <vector>
#include <string>
#include <filesystem>
#include <regex>
#include <unordered_map>
#include <boost/preprocessor/stringize.hpp>

namespace fs = std::filesystem;

using std::vector;
using std::wstring;
using std::make_shared;
using std::shared_ptr;
using std::pair;
using std::make_pair;
using std::string;
using std::unordered_map;
using std::map;

namespace xloil
{
  constexpr const wchar_t* XLOIL_PLUGIN_EXT = L".dll";
  namespace
  {
    static auto emptyTomlTable = toml::table();
  }

  struct LoadedPlugin
  {
    AddinContext* Context;
    HMODULE Handle;
    PluginInitFunc Init;
  };

  static auto& getLoadedPlugins()
  {
    static auto instance = map<
      wstring, LoadedPlugin, CaselessCompare<wchar_t>>();
    return instance;
  }

  bool loadPluginForAddin(AddinContext& context, const std::wstring& pluginName) noexcept
  {
    const auto xllDir = fs::path(context.pathName()).remove_filename();
    const auto coreDir = fs::path(Environment::coreDllPath()).remove_filename();

    auto& loadedPlugins = getLoadedPlugins();
    XLO_INFO("Loading plugins from settings file {}", *context.settings()->source().path);

    // Look for the plugin in the same directory as xloil.dll, 
    // otherwise check the directory of the XLL
    std::error_code fsErr;
    const auto pluginDir = fs::exists(coreDir / (pluginName + XLOIL_PLUGIN_EXT), fsErr)
      ? coreDir
      : xllDir;

    PushDllDirectory setDllDir(pluginDir.c_str());

    const auto pluginPath = pluginDir / (pluginName + XLOIL_PLUGIN_EXT);

    const auto pluginNameUtf8 = utf16ToUtf8(pluginName);

    try
    {
      XLO_INFO(L"Loading plugin {}", pluginName);
        
      const auto pluginSettings = Settings::findPluginSettings(
        context.settings(), pluginNameUtf8.c_str());

      // If the plugin has already be loaded, we just notify it of 
      // a new XLL by calling attach and passing any XLL specific settings
      auto pluginData = loadedPlugins.find(pluginName);
      if (pluginData == loadedPlugins.end())
      {
        auto environment = Settings::environmentVariables(pluginSettings);

        for (auto&[key, val] : environment)
        {
          auto value = expandWindowsRegistryStrings(
            expandEnvironmentStrings(val));
          XLO_DEBUG(L"Setting environment variable: {}='{}'", key, value);

          if (!setEnvironmentVar(key.c_str(), value.c_str()))
            XLO_WARN(L"Failed to set environment variable '{}'");
        }
        // Load the plugin
        const auto lib = LoadLibrary(pluginPath.c_str());
        if (!lib)
          XLO_THROW(writeWindowsError());

        // Find the main entry point required by xlOil
        const auto initFunc = (PluginInitFunc)GetProcAddress(lib,
          BOOST_PP_STRINGIZE(XLO_PLUGIN_INIT_FUNC));
        if (!initFunc)
          XLO_THROW("Couldn't find plugin entry point");


        PluginContext pluginLoadContext =
        {
          PluginContext::Load,
          pluginName.data(),
          pluginSettings ? *pluginSettings.as_table() : emptyTomlTable,
          XLOIL_MAJOR_VERSION,
          XLOIL_MINOR_VERSION,
          XLOIL_PATCH_VERSION
        };
        if (initFunc(&theCoreContext(), pluginLoadContext) < 0)
        {
          //TODO:  Can we roll back any bad registrations?
          FreeLibrary(lib);
          XLO_THROW("Initialisation failed");
        }

        // Add the plugin to the list of loaded plugins
        LoadedPlugin description = { &context, lib, initFunc };
        pluginData = loadedPlugins
          .insert(make_pair(pluginName, description)).first;

        // Register any static functions in the plugin by adding it as a source.
        auto source = make_shared<StaticFunctionSource>(pluginName.c_str());
        context.addSource(source);

        XLO_DEBUG(L"Finished loading plugin {0}", pluginName);
      }

      // Now "attach" the current XLL, passing in its associated settings
      PluginContext pluginAttach =
      {
        PluginContext::Attach,
        pluginName.c_str(),
        pluginSettings ? *pluginSettings.as_table() : emptyTomlTable,
        XLOIL_MAJOR_VERSION,
        XLOIL_MINOR_VERSION,
        XLOIL_PATCH_VERSION
      };
      if (pluginData->second.Init(&context, pluginAttach) < 0)
      {
        XLO_ERROR(L"Failed to attach addin {0} to plugin {1}",
          context.pathName(), pluginName);
        return false;
      }
    }
    catch (const std::exception& e)
    {
      XLO_ERROR(L"Plugin load failed for {0}: {1}\nPath={2}", 
        pluginPath.wstring(), utf8ToUtf16(e.what()), getEnvironmentVar(L"PATH"));
      return false;
    }

    // Undo addition to DLL search path 
    SetDllDirectory(NULL);

    return true;
  }
  
  bool detachPluginForAddin(AddinContext& context, const std::wstring& pluginName) noexcept
  {
    const auto& pluginData = getLoadedPlugins()[pluginName];
    PluginContext pluginDetach =
    {
      PluginContext::Detach,
      pluginName.c_str(),
      emptyTomlTable,
      XLOIL_MAJOR_VERSION,
      XLOIL_MINOR_VERSION,
      XLOIL_PATCH_VERSION
    };
    if (pluginData.Init(&context, pluginDetach) < 0)
    {
      XLO_ERROR(L"Failed to detach addin {0} from plugin {1}",
        context.pathName(), pluginName);
      return false;
    }
    return true;
  }

  bool unloadPluginImpl(const wchar_t* name, LoadedPlugin& plugin) noexcept
  {
    XLO_DEBUG(L"Unloading plugin {0}", name);
    PluginContext context = { 
      PluginContext::Unload, 
      name, 
      emptyTomlTable,          
      XLOIL_MAJOR_VERSION,
      XLOIL_MINOR_VERSION,
      XLOIL_PATCH_VERSION 
    };
    plugin.Init(0, context);
    if (!FreeLibrary(plugin.Handle))
      XLO_WARN(L"FreeLibrary failed for {0}: {1}", name, writeWindowsError());
    return true;
  }
  
  void unloadAllPlugins() noexcept
  {
    for (auto&[name, descr] : getLoadedPlugins())
      unloadPluginImpl(name.c_str(), descr);
    getLoadedPlugins().clear();
  }

  std::vector<std::wstring> listPluginNames()
  {
    std::vector<std::wstring> result;
    std::transform(
      getLoadedPlugins().begin(),
      getLoadedPlugins().end(),
      std::back_inserter(result),
      [](auto it) { return it.first; });
    return std::move(result);
  }

  StaticFunctionSource::StaticFunctionSource(const wchar_t* pluginPath)
    : _sourcePath(pluginPath)
  {}

  void StaticFunctionSource::init()
  {
    auto specs = detail::processRegistryQueue(_sourcePath.c_str());
    registerFuncs(specs);
  }
}
