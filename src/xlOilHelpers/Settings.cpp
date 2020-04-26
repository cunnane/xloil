#include "Settings.h"
#include <xloilHelpers/StringUtils.h>
#include <xloilHelpers/Environment.h>
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
  namespace Settings
  {
    vector<wstring> plugins(const toml::value* root)
    {
      vector<wstring> plugins;
      if (root)
      {
        auto pluginsUtf8 = toml::find_or<vector<string>>(*root, "Plugins", vector<string>());
        std::transform(pluginsUtf8.begin(), pluginsUtf8.end(),
          std::back_inserter(plugins), utf8ToUtf16);
      }
      return plugins;
    }

    namespace
    {
      std::string findStr(const toml::value* root, const char* tag, const char* default)
      {
        return root
          ? toml::find_or<string>(*root, tag, default)
          : default;
      }
    }
    std::wstring pluginSearchPattern(const toml::value* root)
    {
      return utf8ToUtf16(findStr(root, "PluginSearchPattern", ""));
    }
    std::wstring logFilePath(const toml::value* root)
    {
      return utf8ToUtf16(findStr(root, "LogFile", ""));
    }
    std::string logLevel(const toml::value* root)
    {
      return findStr(root, "LogLevel", "warn");
    }
  }
  std::shared_ptr<const toml::value> findSettingsFile(const wchar_t* dllPath)
  {
    fs::path path;
 
    auto settingsFileName = fs::path(dllPath).filename().replace_extension(XLOIL_SETTINGS_FILE_EXT);

    // First check the same directory as the dll itself
    path = fs::path(dllPath).remove_filename() / settingsFileName;
      
    // Then look in the user's appdata
    if (!fs::exists(path))
      path = fs::path(getEnvVar(L"APPDATA")) / L"xlOil" / settingsFileName;

    if (fs::exists(path))
      return make_shared<toml::value>(toml::parse(path.string()));
    
    return shared_ptr<const toml::value>();
  }
}