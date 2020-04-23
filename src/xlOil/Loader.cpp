#include "Loader.h"
#include "WindowsSlim.h"
#include "Interface.h"
#include "Log.h"
#include "Settings.h"
#include "EntryPoint.h"
#include "internal/FuncRegistry.h"
#include <xlOil/Loaders/AddinLoader.h>
#include <toml11/toml.hpp>
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

namespace
{

  // TODO: support hives other than HKLM!

  wstring getRegistryValue(wstring location)
  {
    const auto lastSlash = location.rfind(L'\\');
    const auto subKey = location.substr(0, lastSlash);
    const auto value = lastSlash + 1 < location.size() ? location.substr(lastSlash + 1) : wstring();
    wchar_t buffer[1024];
    DWORD bufSize = sizeof(buffer) / sizeof(wchar_t);
    if (ERROR_SUCCESS == RegGetValue(
      HKEY_LOCAL_MACHINE,
      subKey.c_str(),
      value.c_str(),
      RRF_RT_REG_SZ,
      nullptr /*type not required*/,
      &buffer,
      &bufSize))
    {
      return wstring(buffer, buffer + bufSize);
    }
    XLO_THROW(L"Registry key HKLM\\{0} missing when reading settings file", location);
  }
}

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

    // Should match "<HKLM\(Reg\Key\Value)>", used for expanding 
    // environment variable strings later
    std::wregex registryExpander(L"<HKLM\\\\([^>]*)>", 
      std::regex_constants::optimize | std::regex_constants::ECMAScript);

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
        
        toml::value pluginSettings;
        if (context->settings())
          pluginSettings = toml::find_or<toml::table>(
            *context->settings(), utf16ToUtf8(pluginName), toml::table());

        // If the plugin has already be loaded, we just notify it of 
        // a new XLL by calling attach and passing any XLL specific settings
        auto found = getLoadedPlugins().find(pluginName);
        if (found != getLoadedPlugins().end())
        {
          PluginContext plugin = { PluginContext::Attach, pluginName.c_str(), &pluginSettings };
          if (found->second.Init(context, plugin) < 0)
            XLO_ERROR(L"Failed to attach addin {0} to plugin {1}", context->pathName(), pluginName);
        }
        else
        {
          auto environment = toml::find_or<toml::table>(
            pluginSettings, "Environment", toml::table());

          // Settings in the enviroment block looks like key=val
          // We interpret this as an environment variable to set
          for (auto[key, val] : environment)
          {
            wstring value = expandEnvVars(utf8ToUtf16(val.as_string().str).c_str());
            std::wsmatch match;
            std::regex_match(value, match, registryExpander);
            if (match.size() == 2)
              value = getRegistryValue(match[1].str());
             
            environmentVariables.emplace_back(make_shared<PushEnvVar>(
              utf8ToUtf16(key).c_str(),
              value.c_str()));
          }

          // Load the plugin
          auto lib = LoadLibrary(path.c_str());
          if (!lib)
            XLO_THROW(writeWindowsError());

          // Find the main entry point required by xlOil
          auto initFunc = (PluginInitFunc)GetProcAddress(lib, 
            BOOST_PP_STRINGIZE(XLO_PLUGIN_INIT_FUNC));
          if (!initFunc)
            XLO_THROW("Couldn't find plugin entry point");

          // TODO: check build key xloil_buildId for version control
          //if ( != 0)  
          //{
          //  FreeLibrary(lib);
          //  continue;
          //}

          // First load the plugin using any settings that have been specified in the
          // core config file
          toml::value coreSettings = theCoreContext()->settings()
            ? toml::find_or<toml::table>(
                *theCoreContext()->settings(), utf16ToUtf8(pluginName), toml::table())
            : toml::value();

          PluginContext pluginCore = { PluginContext::Load, pluginName.c_str(), &coreSettings };
          if (initFunc(theCoreContext(), pluginCore) < 0)
          {
            //TODO:  Can we roll back any bad registrations?
            FreeLibrary(lib);
            XLO_THROW("Initialisation failed");
          }

          // Now "attach" the current XLL, passing in its associated settings
          PluginContext plugin = { PluginContext::Attach, pluginName.c_str(), &pluginSettings };
          initFunc(context, plugin);

          // Add the plugin to the list of loaded plugins
          LoadedPlugin description = { context, lib, initFunc };
          getLoadedPlugins().insert(make_pair(pluginName, description));

          // Register any static functions in the plugin by adding
          // it as a source.
          context->tryAdd<StaticFunctionSource>(pluginName.c_str(), pluginName.c_str());
        }
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Plugin load failed for {0}: {1}", path.string(), e.what());
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
