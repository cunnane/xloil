#include "Settings.h"
#include "EntryPoint.h"
#include "Interface.h"
#include "StringUtils.h"
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
using std::shared_ptr;
using std::make_shared;

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

        auto root = findSettingsFile(theXllPath());
        if (!root)
          root = make_shared<toml::value>();

        logFilePath = toml::find_or<string>(*root, "LogFile", "");
        logLevel = toml::find_or<string>(*root, "LogLevel", "warn");
        pluginSearchPattern = toml::find_or<string>(*root, "PluginSearchPattern", "");
        auto pluginsUtf8 = toml::find_or<vector<string>>(*root, "Plugins", vector<string>());
        std::transform(pluginsUtf8.begin(), pluginsUtf8.end(), std::back_inserter(plugins), utf8ToUtf16);
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

  std::shared_ptr<toml::value> findSettingsFile(const wchar_t* dllPath)
  {
    fs::path path;
 
    // First look in the user's appdata
    auto settingsFileName = fs::path(dllPath).filename().replace_extension(XLOIL_SETTINGS_FILE_EXT);
    path = fs::path(getEnvVar(L"%APPDATA%")) / L"xlOil" / settingsFileName;
      
    // Then check the same directory as the dll itself
    if (!fs::exists(path))
      path = fs::path(dllPath).remove_filename() / settingsFileName;

    if (fs::exists(path))
      return make_shared<toml::value>(toml::parse(path.string()));

    return shared_ptr<toml::value>();
  }
}