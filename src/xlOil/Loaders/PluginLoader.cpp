#include "PluginLoader.h"
#include <xlOilHelpers/WindowsSlim.h>
#include <xlOilHelpers/Environment.h>
#include <COMInterface/RtdManager.h>
#include <xlOil/Interface.h>
#include <xlOil/Log.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/Register/FuncRegistry.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <tomlplusplus/toml.hpp>
#include <vector>
#include <string>
#include <filesystem>
#include <regex>
#include <set>
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
  struct LoadedPlugin
  {
    AddinContext* Context;
    HMODULE Handle;
    PluginInitFunc Init;
  };

  template <class TChar>
  struct CaselessCompare 
  {
    bool operator()(
      const std::basic_string<TChar> & lhs, 
      const std::basic_string<TChar> & rhs) const 
    {
      return (*this)(lhs.c_str(), rhs.c_str());
    }
    bool operator()(const TChar* lhs, const TChar* rhs) const 
    {
      if constexpr (std::is_same<TChar, wchar_t>::value)
        return _wcsicmp(lhs, rhs) < 0;
      else
        return _stricmp(lhs, rhs) < 0;
    }
  };

  static auto& getLoadedPlugins()
  {
    static auto instance = map<
      wstring, LoadedPlugin, CaselessCompare<wchar_t>>();
    return instance;
  }

  void loadPlugins(
    AddinContext* context, 
    const std::vector<std::wstring>& names) noexcept
  {
    auto plugins = std::set<wstring>(names.cbegin(), names.cend());

    const auto xllDir = fs::path(context->pathName()).remove_filename();
    const auto coreDir = fs::path(Core::theCorePath()).remove_filename();

    // If the settings specify a search pattern for plugins, 
    // find the DLLs and add them to our plugins collection
    if (context->settings())
    {
      WIN32_FIND_DATA fileData;

      auto searchPath = xllDir / Settings::pluginSearchPattern(
        (*context->settings())["Addin"]);
      auto fileHandle = FindFirstFile(searchPath.c_str(), &fileData);
      if (fileHandle != INVALID_HANDLE_VALUE &&
        fileHandle != (void*)ERROR_FILE_NOT_FOUND)
      {
        do
        {
          if (_wcsicmp(fileData.cFileName, Core::theCoreName()) == 0)
            continue;

          plugins.emplace(fs::path(fileData.cFileName).stem()); // TODO: remove extn?
        } while (FindNextFile(fileHandle, &fileData));
      }
    }

    
    auto& loadedPlugins = getLoadedPlugins();

    for (const auto& pluginName : plugins)
    {
      // Look for the plugin in the same directory as xloil.dll, 
      // otherwise check the directory of the XLL
      const auto pluginDir = fs::exists(coreDir / (pluginName + L".dll"))
        ? coreDir
        : xllDir;

      SetDllDirectory(pluginDir.c_str());

      const auto pluginPath = pluginDir / (pluginName + L".dll");

      const auto pluginNameUtf8 = utf16ToUtf8(pluginName);

      try
      {
        XLO_INFO(L"Loading plugin {}", pluginName);
        
        const auto pluginSettings = Settings::findPluginSettings(
          context->settings(), pluginNameUtf8.c_str());

        // If the plugin has already be loaded, we just notify it of 
        // a new XLL by calling attach and passing any XLL specific settings
        auto pluginData = loadedPlugins.find(pluginName);
        if (pluginData == loadedPlugins.end())
        {
          // First load the plugin using any settings that have been specified in the
          // core config file, otherwise the ones in the add-ins ini file. This avoids
          // race conditions with different add-in load orders.

          auto loadSettings = theCoreContext()->settings()
            ? Settings::findPluginSettings(theCoreContext()->settings(), pluginNameUtf8.c_str())
            : pluginSettings;

          auto environment = Settings::environmentVariables(loadSettings);

          for (auto&[key, val] : environment)
          {
            auto value = expandWindowsRegistryStrings(
              expandEnvironmentStrings(val));

            SetEnvironmentVariable(key.c_str(), value.c_str());
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

          // TODO: check build key xloil_buildId for version control

          PluginContext pluginLoadContext =
          {
            PluginContext::Load,
            pluginName.c_str(),
            loadSettings.as_table()
          };
          if (initFunc(theCoreContext(), pluginLoadContext) < 0)
          {
            //TODO:  Can we roll back any bad registrations?
            FreeLibrary(lib);
            XLO_THROW("Initialisation failed");
          }

          // Add the plugin to the list of loaded plugins
          LoadedPlugin description = { context, lib, initFunc };
          pluginData = loadedPlugins
            .insert(make_pair(pluginName, description)).first;

          // Register any static functions in the plugin by adding
          // it as a source.
          context->tryAdd<StaticFunctionSource>(pluginName.c_str(), pluginName.c_str());
        }

        // Now "attach" the current XLL, passing in its associated settings
        PluginContext pluginAttach = 
        { 
          PluginContext::Attach, 
          pluginName.c_str(), 
          pluginSettings.as_table()
        };
        if (pluginData->second.Init(context, pluginAttach) < 0)
          XLO_ERROR(L"Failed to attach addin {0} to plugin {1}", 
            context->pathName(), pluginName);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(L"Plugin load failed for {0}: {1}\nPath={2}", 
          pluginPath.wstring(), utf8ToUtf16(e.what()), getEnvVar(L"PATH"));
      }

      // Undo addition to DLL search path 
      SetDllDirectory(NULL);
    }
  }

  bool unloadPluginImpl(const wchar_t* name, LoadedPlugin& plugin) noexcept
  {
    XLO_DEBUG(L"Unloading plugin {0}", name);
    PluginContext context = { PluginContext::Unload, name, nullptr };
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
}
