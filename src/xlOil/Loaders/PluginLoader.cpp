#include "PluginLoader.h"
#include <xlOilHelpers/WindowsSlim.h>
#include <xlOilHelpers/Environment.h>
#include <xlOil/Interface.h>
#include <xlOil/Log.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/EntryPoint.h>
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

namespace xloil
{
  struct LoadedPlugin
  {
    AddinContext* Context;
    HMODULE Handle;
    PluginInitFunc Init;
  };

  static auto& getLoadedPlugins()
  {
    static auto instance = unordered_map<wstring, LoadedPlugin>();
    return instance;
  }

  void loadPlugins(AddinContext* context, const std::vector<std::wstring>& names) noexcept
  {
    auto plugins = std::set<wstring>(names.cbegin(), names.cend());

    const auto xllDir = fs::path(context->pathName()).remove_filename();
    const auto coreDir = fs::path(Core::theCorePath()).remove_filename();

    // If the settings specify a search pattern for plugins, 
    // find the DLLs and add them to our plugins collection
    if (context->settings())
    {
      WIN32_FIND_DATA fileData;

      auto searchPath = xllDir / Settings::pluginSearchPattern(context->settings());
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

    for (auto pluginName : plugins)
    {
      // Look for the plugin in the same directory as xloil.dll, 
      // otherwise check the directory of the XLL
      const auto pluginDir = fs::exists(coreDir / (pluginName + L".dll"))
        ? coreDir
        : xllDir;

      SetDllDirectory(pluginDir.c_str());

      const auto path = pluginDir / (pluginName + L".dll");

      try
      {
        // The PushEnvVar class will remove any set environment
        // variables when it goes out of scope
        vector<shared_ptr<PushEnvVar>> environmentVariables;

        XLO_INFO(L"Found plugin {}", pluginName);
        
        toml::node_view<const toml::node> pluginSettings;
        if (context->settings())
          pluginSettings = (*context->settings())[utf16ToUtf8(pluginName)];

        // If the plugin has already be loaded, we just notify it of 
        // a new XLL by calling attach and passing any XLL specific settings
        auto pluginData = getLoadedPlugins().find(pluginName);
        if (pluginData == getLoadedPlugins().end())
        {
          // First load the plugin using any settings that have been specified in the
          // core config file, otherwise the ones in the add-ins ini file. This avoids
          // race conditions with differnt add-in load orders.
          auto loadSettings = theCoreContext()->settings()
            ? (*theCoreContext()->settings())[utf16ToUtf8(pluginName)]
            : pluginSettings;

          auto environment = loadSettings["Environment"].as_array();

          // Settings in the enviroment block looks like key=val
          // We interpret this as an environment variable to set
          if (environment)
            for (auto& innerTable : *environment)
            {
              for (auto[key, val] : *innerTable.as_table())
              {
                auto value = expandWindowsRegistryStrings(
                  expandEnvironmentStrings(
                    utf8ToUtf16(val.value_or(""))));

                environmentVariables.emplace_back(std::make_shared<PushEnvVar>(
                  utf8ToUtf16(key).c_str(),
                  value.c_str()));
              }
            }

          // Load the plugin
          const auto lib = LoadLibrary(path.c_str());
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
          pluginData = getLoadedPlugins()
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
          path.wstring(), utf8ToUtf16(e.what()), getEnvVar(L"PATH"));
      }

      // Undo addition to DLL search path 
      SetDllDirectory(NULL);
    }
  }

  void unloadPlugins() noexcept
  {
    for (auto[name, descr] : getLoadedPlugins())
    {
      XLO_DEBUG(L"Unloading plugin {0}", name);
      PluginContext plugin = { PluginContext::Unload, name.c_str(), nullptr };
      descr.Init(0, plugin);
      if (!FreeLibrary(descr.Handle))
        XLO_WARN(L"FreeLibrary failed for {0}: {1}", name, writeWindowsError());
    }
    getLoadedPlugins().clear();
  }
}
