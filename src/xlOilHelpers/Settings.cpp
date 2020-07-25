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
      auto findStr(const toml::view_node& root, const char* tag, const char* default)
      {
        return root[tag].value_or<string>(default);
      }
      auto findVecStr(const toml::view_node& root, const char* tag)
      {
        vector<wstring> result;
        auto utf8 = root[tag].as_array();
        if (utf8)
          for (auto& x : *utf8)
            result.push_back(utf8ToUtf16(x.value_or("")));
        return result;
      }
    }
    vector<wstring> plugins(const toml::view_node& root)
    {
      return findVecStr(root, "Plugins");
    }
    std::wstring pluginSearchPattern(const toml::view_node& root)
    {
      return utf8ToUtf16(findStr(root, "PluginSearchPattern", ""));
    }
    std::wstring logFilePath(const toml::view_node& root)
    {
      return utf8ToUtf16(findStr(root, "LogFile", ""));
    }
    std::string logLevel(const toml::view_node& root)
    {
      return findStr(root, "LogLevel", "warn");
    }
    std::vector<std::wstring> dateFormats(const toml::view_node& root)
    {
      return findVecStr(root, "DateFormats");
    }
    std::vector<std::pair<std::wstring, std::wstring>> 
      environmentVariables(const toml::view_node& root)
    {
      vector<pair<wstring, wstring>> result;
      auto environment = root["Environment"].as_array();
      if (environment)
        for (auto& innerTable : *environment)
        {
          // Settings in the enviroment block looks like key=val
          // We interpret this as an environment variable to set
          for (auto[key, val] : *innerTable.as_table())
          {
            result.emplace_back(make_pair(
              utf8ToUtf16(key),
              utf8ToUtf16(val.value_or(""))));
          }
        }
      return result;
    }
    toml::node_view<const toml::node> findPluginSettings(
      const toml::table* table, const char* name)
    {
      // Note: if you get a compile error here, make sure the ctor for node_view
      // is public.  It's hidden in the original code, which means it can only
      // be constructed from a table, not a table iterator. This is an inconvenience
      // which I fixed!
      if (table)
        for (auto i = table->cbegin(); i != table->cend(); ++i)
        {
          if (_stricmp((*i).key.c_str(), name) == 0)
            return &(*i).value;
        }
   
      return toml::node_view<const toml::node>();
    }
  }
  std::shared_ptr<const toml::table> findSettingsFile(const wchar_t* dllPath)
  {
    fs::path path;
 
    const auto settingsFileName = 
      fs::path(dllPath).filename().replace_extension(XLOIL_SETTINGS_FILE_EXT);
    
    // Look in the user's appdata
    path = fs::path(getEnvVar(L"APPDATA")) / L"xlOil" / settingsFileName;

    // Then check the same directory as the dll itself
    if (!fs::exists(path))
      path = fs::path(dllPath).remove_filename() / settingsFileName;

    return fs::exists(path) 
      ? make_shared<toml::table>(toml::parse_file(path.string())) 
      : shared_ptr<const toml::table>();
  }
}