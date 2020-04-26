#pragma once
#include <vector>
#include <string>
#include <unordered_map>

namespace toml {
  template<typename, template<typename...> class, template<typename...> class> 
    class basic_value;
  struct discard_comments;
  using value = basic_value<discard_comments, std::unordered_map, std::vector>;
}

namespace xloil
{
  constexpr char* XLOIL_SETTINGS_FILE_EXT = "ini";

  namespace Settings
  {
    std::wstring logFilePath(const toml::value* root);
    std::string logLevel(const toml::value* root);
    std::vector<std::wstring> plugins(const toml::value* root);
    std::wstring pluginSearchPattern(const toml::value* root);
  };

  std::shared_ptr<const toml::value>
    findSettingsFile(const wchar_t* dllPath);
}