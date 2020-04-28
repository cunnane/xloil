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
   

    namespace
    {
      auto findStr(const toml::value* root, const char* tag, const char* default)
      {
        return root
          ? toml::find_or<string>(*root, tag, default)
          : default;
      }
      auto findVecStr(const toml::value* root, const char* tag)
      {
        vector<wstring> result;
        if (root)
        {
          auto utf8 = toml::find_or<vector<string>>(*root, tag, vector<string>());
          std::transform(utf8.begin(), utf8.end(), 
            std::back_inserter(result), utf8ToUtf16);
        }
        return result;
      }
    }
    vector<wstring> plugins(const toml::value* root)
    {
      return findVecStr(root, "Plugins");
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
    std::vector<std::wstring> dateFormats(const toml::value* root)
    {
      return findVecStr(root, "DateFormats");
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