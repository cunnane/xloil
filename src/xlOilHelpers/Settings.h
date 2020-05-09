#pragma once
#include <vector>
#include <string>
#include <unordered_map>

namespace toml {
  class table;
}

namespace xloil
{
  constexpr char* XLOIL_SETTINGS_FILE_EXT = "ini";

  namespace Settings
  {
    std::wstring logFilePath(const toml::table* root);
    std::string logLevel(const toml::table* root);
    std::vector<std::wstring> plugins(const toml::table* root);
    std::wstring pluginSearchPattern(const toml::table* root);
    std::vector<std::wstring> dateFormats(const toml::table* root);
  };

  std::shared_ptr<const toml::table>
    findSettingsFile(const wchar_t* dllPath);
}