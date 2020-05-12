#pragma once
#include <vector>
#include <string>
#include <unordered_map>

namespace toml {
  class node; class table;
  template <typename T> class node_view;

  using view_node = toml::node_view<const toml::node>;
}

namespace xloil
{
  constexpr char* XLOIL_SETTINGS_FILE_EXT = "ini";

  namespace Settings
  {
    std::wstring logFilePath(const toml::view_node& root);
    std::string logLevel(const toml::view_node& root);
    std::vector<std::wstring> plugins(const toml::view_node& root);
    std::wstring pluginSearchPattern(const toml::view_node& root);
    std::vector<std::wstring> dateFormats(const toml::view_node& root);
    std::vector<std::pair<std::wstring, std::wstring>>
      environmentVariables(const toml::view_node& root);
  };

  std::shared_ptr<const toml::table>
    findSettingsFile(const wchar_t* dllPath);
}