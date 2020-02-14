#include "Loader.h"
#include "WindowsSlim.h"
#include "Interface.h"
#include "Log.h"
#include "Register.h"
#include "Settings.h"
#include "internal/FuncRegistry.h"
#include <toml11/toml.hpp>
#include <vector>
#include <string>
#include <filesystem>
#include <regex>

namespace fs = std::filesystem;

using std::vector;
using std::wstring;
using std::make_shared;
using std::shared_ptr;
using std::pair;
using std::make_pair;
using std::string;

namespace
{
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
    XLO_THROW(L"HKLM\\{0} not found", location);
  }
}

namespace xloil
{
  static auto& getLoadedPlugins()
  {
    static auto instance = vector<pair<HMODULE, shared_ptr<Core>>>();
    return instance;
  }

  void loadPlugins() noexcept
  {
    processRegistryQueue(Core::theCoreName());

    auto& plugins = theCoreSettings().pluginNamesAndPath;

    auto corePath = fs::path(Core::theCorePath()).remove_filename();

    // The data for each file we find.
    WIN32_FIND_DATA fileData;

    // Find the DLLs in the plugin folder and add them to our plugins
    auto searchPath = corePath / theCoreSettings().pluginSearchPattern;
    auto fileHandle = FindFirstFile(searchPath.c_str(), &fileData);
    if (fileHandle != INVALID_HANDLE_VALUE &&
      fileHandle != (void*)ERROR_FILE_NOT_FOUND)
    {
      do
      {
        if (_wcsicmp(fileData.cFileName, Core::theCoreName()) == 0)
          continue;
        if (_wcsicmp(fileData.cFileName, L"xlOil_Loader.dll") == 0)
          continue;
        // Check we don't already have this filename (2nd pair item)
        if (std::none_of(plugins.begin(), plugins.end(),
          [fileData](auto x) { return _wcsicmp(fileData.cFileName, x.second.c_str()) == 0; }))
        {
          plugins.emplace_back(fileData.cFileName, fileData.cFileName);
        }
      } while (FindNextFile(fileHandle, &fileData));
    }

    SetDllDirectory(corePath.c_str());

    // Should match "<HKLM\(Reg\Key\Value)>"
    std::wregex registryExpander(L"<HKLM\\\\([^>]*)>", std::regex_constants::optimize);

    for (auto[pluginName, pluginPath] : plugins)
    {
      vector<shared_ptr<PushEnvVar>> pathPusher;

      XLO_TRACE(L"Found plugin {}", pluginName);
      auto path = fs::path(pluginPath);
      if (path.is_relative())
        path = corePath / path;

      auto settings = fetchPluginSettings(pluginName.c_str());
      if (settings)
      {
        auto environment = toml::find_or<toml::table>(*settings, "Environment", toml::table());
        for (auto[key, val] : environment)
        {
          wstring value = utf8_to_wstring(val.as_string().str);
          std::wsmatch match;
          std::regex_match(value, match, registryExpander);
          if (match.size() == 2)
            value = getRegistryValue(match[1].str());

          pathPusher.emplace_back(make_shared<PushEnvVar>(
            utf8_to_wstring(key).c_str(),
            value.c_str()));
        }
      }

      // Load the plugin
      auto lib = LoadLibrary(path.c_str());
      if (!lib)
      {
        auto err = writeWindowsError();
        XLO_WARN(L"Couldn't load plugin at {0}: {1}", path.c_str(), err);
        continue;
      }

      auto initFunc = (pluginInitFunc)GetProcAddress(lib, XLO_STR(XLO_PLUGIN_INIT_FUNC));
      if (!initFunc)
      {
        XLO_WARN(L"Couldn't find entry point for plugin {0}", pluginPath);
        continue;
      }

      // TODO: check build key xloil_buildId for version control
      //if ( != 0)  
      //{
      //  FreeLibrary(lib);
      //  continue;
      //}
      
      auto coreObj = make_shared<Core>(pluginName.c_str());
      if (initFunc(*coreObj) < 0)
      {
        // TODO:  Can we roll back any bad registrations?
        XLO_ERROR(L"Plugin initialisation failed for {}", pluginPath);
        FreeLibrary(lib);
        continue;
      }

      getLoadedPlugins().emplace_back(lib, coreObj);
    } 

    // Undo addition to DLL search path 
    SetDllDirectory(NULL);
  }

  void unloadPlugins() noexcept
  {
    for (auto& m : getLoadedPlugins())
    {
      XLO_TRACE(L"Unloading plugin {0}", m.second->pluginName());
      auto exitFunc = (pluginExitFunc)GetProcAddress(m.first, XLO_STR(XLO_PLUGIN_EXIT_FUNC));
      if (exitFunc)
        exitFunc();
      //m.second->deregisterAll();
      if (!FreeLibrary(m.first))
        XLO_WARN(L"FreeLibrary failed for {0}: {1}", m.second->pluginName(), writeWindowsError());
    }
    getLoadedPlugins().clear();
  }
}
