#include "Settings.h"
#include <xloilHelpers/StringUtils.h>
#include <xloilHelpers/Environment.h>
#include <tomlplusplus/toml.hpp>
#include <filesystem>
#include <fstream>

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
      auto findStr(const toml::table* root, const char* tag, const char* default)
      {
        return root
          ? (*root)[tag].value_or<string>(default)
          : default;
      }
      auto findVecStr(const toml::table* root, const char* tag)
      {
        vector<wstring> result;
        if (root)
        {
          auto utf8 = (*root)[tag].as_array();
          if (utf8)
            for (auto& x : *utf8)
              result.push_back(utf8ToUtf16(x.value_or("")));
        }
        return result;
      }
    }
    vector<wstring> plugins(const toml::table* root)
    {
      return findVecStr(root, "Plugins");
    }
    std::wstring pluginSearchPattern(const toml::table* root)
    {
      return utf8ToUtf16(findStr(root, "PluginSearchPattern", ""));
    }
    std::wstring logFilePath(const toml::table* root)
    {
      return utf8ToUtf16(findStr(root, "LogFile", ""));
    }
    std::string logLevel(const toml::table* root)
    {
      return findStr(root, "LogLevel", "warn");
    }
    std::vector<std::wstring> dateFormats(const toml::table* root)
    {
      return findVecStr(root, "DateFormats");
    }
  }
  std::shared_ptr<const toml::table> findSettingsFile(const wchar_t* dllPath)
  {
    fs::path path;
 
    auto settingsFileName = fs::path(dllPath).filename().replace_extension(XLOIL_SETTINGS_FILE_EXT);

    // First check the same directory as the dll itself
    path = fs::path(dllPath).remove_filename() / settingsFileName;
      
    // Then look in the user's appdata
    if (!fs::exists(path))
      path = fs::path(getEnvVar(L"APPDATA")) / L"xlOil" / settingsFileName;

    if (fs::exists(path))
      return make_shared<toml::table>(toml::parse_file(path.string()));
    
    return shared_ptr<const toml::table>();
  }
}