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

        logFilePath = toml::find_or<string>(root, "LogFile", "");
        logLevel = toml::find_or<string>(root, "LogLevel", "warn");
        pluginSearchPattern = toml::find_or<string>(root, "PluginSearchPattern", "");
        plugins = utf8_to_wstring_v(toml::find_or<vector<string>>(root, "Plugins", vector<string>()));
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
}