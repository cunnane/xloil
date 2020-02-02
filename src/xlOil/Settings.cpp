#include "Settings.h"
#include "EntryPoint.h"
#include "Interface.h"
#include "Utils.h"
#include "Log.h"
#include "Options.h"
#include <toml11/toml.hpp>
#include <filesystem>

namespace fs = std::filesystem;

using std::vector;
using std::string;
using std::wstring;
using std::pair;
using std::make_pair;

namespace xloil
{

  class SettingsReader : public Settings
  {
  public:
    static SettingsReader& get()
    {
      static SettingsReader instance;
      return instance;
    }
    
    const auto& pluginSettings() const { return _pluginSettings; }

  private:
    SettingsReader()
    {
      try
      {
        auto corePath = fs::path(theXllPath()).replace_extension(XLOIL_SETTINGS_FILE_EXT);
        const toml::value root = toml::parse(corePath.string());

        // Process core settings
        auto& core = toml::find(root, "Core");

        logFilePath = toml::find_or<std::string>(core, "LogFile", "");
        logLevel = toml::find_or<std::string>(core, "LogLevel", "warn");

        // Process plugin settings
        auto& plugins = toml::find<toml::table>(root, "Plugins");
        for (auto i : plugins)
        {
          if (i.second.is_table())
            _pluginSettings.insert(make_pair(i.first, i.second));
        }

        for (auto[key, val] : _pluginSettings)
        {
          auto& table = val.as_table();
          auto ipath = table.find("PluginPath");
          auto path = ipath != table.end() 
            ? ipath->second.as_string() 
            : (key + ".dll");
          pluginNamesAndPath.emplace_back(
            utf8_to_wstring(key.c_str()),
            utf8_to_wstring(path.str.c_str()));
        }
      }
      catch (const std::exception& e)
      {
        // TODO: obviously the logger won't be properly setup...ideas?
        XLO_ERROR("Error processing settings file: {}", e.what());
      }
    }

    std::unordered_map<string, toml::value> _pluginSettings;
  };

  Settings& theCoreSettings()
  {
    return SettingsReader::get();
  }

  const toml::value* fetchPluginSettings(const wchar_t* pluginName)
  {
    auto& pluginSettings = SettingsReader::get().pluginSettings();
    auto i = pluginSettings.find(wstring_to_utf8(pluginName));
    if (i != pluginSettings.end())
      return &i->second;
    
    return nullptr;
  }

}