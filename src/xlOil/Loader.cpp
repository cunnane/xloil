#include "Loader.h"
#include "WindowsSlim.h"
#include "Interface.h"
#include "Log.h"
#include "Register.h"
#include "Settings.h"
#include "FuncRegistry.h"
#include <toml11/toml.hpp>
#include <vector>
#include <string>
#include <filesystem>
#include <boost/preprocessor/stringize.hpp>
#include <filesystem>
namespace fs = std::filesystem;

using std::vector;
using std::wstring;
using std::make_shared;
using std::shared_ptr;
using std::pair;
using std::make_pair;
using std::string;


namespace xloil
{
  // TODO: any way to pick this up automatically?
#define XLO_CORE_DLL_NAME "xlOil_Core.dll"
#define XLO_CORE_DLL_NAMEW L"xlOil_Core.dll"


  static constexpr wchar_t* thePluginPrefix = L"xloil_";

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
    auto searchPath = corePath / (wstring(thePluginPrefix) + L"*.dll");
    auto fileHandle = FindFirstFile(searchPath.c_str(), &fileData);
    if (fileHandle != INVALID_HANDLE_VALUE &&
      fileHandle != (void*)ERROR_FILE_NOT_FOUND)
    {
      do
      {
        if (_wcsicmp(fileData.cFileName, XLO_CORE_DLL_NAMEW) == 0)
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

    // Pair is (plugin_name, file_path_
    for (auto plug : plugins)
    {
      vector<shared_ptr<PushEnvVar>> pathPusher;

      XLO_TRACE(L"Found plugin {}", plug.first);
      auto path = fs::path(plug.second);
      if (path.is_relative())
        path = corePath / path;

      auto settings = fetchPluginSettings(plug.first.c_str());
      if (settings)
      {
        auto environment = toml::find_or<toml::table>(*settings, "Environment", toml::table());
        for (auto var : environment)
        {
          pathPusher.emplace_back(make_shared<PushEnvVar>(
            utf8_to_wstring(var.first).c_str(), 
            utf8_to_wstring((string)var.second.as_string()).c_str()));
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

      auto initFunc = (pluginInitFunc)GetProcAddress(lib, BOOST_PP_STRINGIZE(XLO_PLUGIN_INIT_FUNC));
      if (!initFunc)
      {
        XLO_WARN(L"Couldn't find entry point for plugin {0}", plug.second);
        continue;
      }

      // TODO: check build key xloil_buildId for version control
      //if ( != 0)  
      //{
      //  FreeLibrary(lib);
      //  continue;
      //}
      
      auto coreObj = make_shared<Core>(plug.first.c_str());
      if (initFunc(*coreObj) < 0)
      {
        // TODO:  Can we roll back any bad registrations?
        XLO_ERROR(L"Plugin initialisation failed for {}", plug.second);
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
      auto exitFunc = (pluginExitFunc)GetProcAddress(m.first, BOOST_PP_STRINGIZE(XLO_PLUGIN_EXIT_FUNC));
      if (exitFunc)
        exitFunc();
      //m.second->deregisterAll();
      if (!FreeLibrary(m.first))
        XLO_WARN(L"FreeLibrary failed for {0}: {1}", m.second->pluginName(), writeWindowsError());
    }
    getLoadedPlugins().clear();
  }
}
