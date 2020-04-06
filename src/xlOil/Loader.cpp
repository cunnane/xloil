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
#include <set>

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
    XLO_THROW(L"Registry key HKLM\\{0} missing when reading settings file", location);
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

    auto plugins = std::set<wstring>(
      theCoreSettings().plugins.cbegin(),
      theCoreSettings().plugins.cend());

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

        plugins.emplace(fileData.cFileName);
      } while (FindNextFile(fileHandle, &fileData));
    }

    SetDllDirectory(corePath.c_str());

    // Should match "<HKLM\(Reg\Key\Value)>"
    std::wregex registryExpander(L"<HKLM\\\\([^>]*)>", std::regex_constants::optimize);

    for (auto pluginName : plugins)
    {
      const auto path = corePath / pluginName;

      try
      {
        vector<shared_ptr<PushEnvVar>> pathPusher;

        XLO_INFO(L"Found plugin {}", pluginName);
        auto settings = findSettingsFile(path.c_str());
        if (settings)
        {
          auto environment = toml::find_or<toml::table>(*settings, "Environment", toml::table());
          for (auto[key, val] : environment)
          {
            wstring value = utf8ToUtf16(val.as_string().str);
            std::wsmatch match;
            std::regex_match(value, match, registryExpander);
            if (match.size() == 2)
              value = getRegistryValue(match[1].str());

            pathPusher.emplace_back(make_shared<PushEnvVar>(
              utf8ToUtf16(key).c_str(),
              value.c_str()));
          }
        }

        // Load the plugin
        auto lib = LoadLibrary(path.c_str());
        if (!lib)
          XLO_THROW(writeWindowsError());

        auto initFunc = (pluginInitFunc)GetProcAddress(lib, XLO_STR(XLO_PLUGIN_INIT_FUNC));
        if (!initFunc)
          XLO_THROW("Couldn't find plugin entry point");

        // TODO: check build key xloil_buildId for version control
        //if ( != 0)  
        //{
        //  FreeLibrary(lib);
        //  continue;
        //}

        auto coreObj = make_shared<Core>(pluginName.c_str(), settings);
        if (initFunc(*coreObj) < 0)
        {
          // TODO:  Can we roll back any bad registrations?
          FreeLibrary(lib);
          XLO_THROW("Initialisation failed");
        }

        getLoadedPlugins().emplace_back(lib, coreObj);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Plugin load failed for {0}: {1}", path.string(), e.what());
      }
    }
    // Undo addition to DLL search path 
    SetDllDirectory(NULL);
  }

  void unloadPlugins() noexcept
  {
    for (auto& m : getLoadedPlugins())
    {
      XLO_DEBUG(L"Unloading plugin {0}", m.second->pluginName());
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
