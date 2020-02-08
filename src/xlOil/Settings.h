#pragma once
#include <vector>
#include <string>
#include <unordered_map>

// TODO: using the fwd declare gives a link error. Posssible to fix?
//namespace toml { class value; }
//#include <toml11/toml.hpp>

namespace toml {
  template<typename, template<typename...> class, template<typename...> class> class basic_value;
  struct discard_comments;
  using value = basic_value<discard_comments, std::unordered_map, std::vector>;
}

namespace xloil
{
  struct Settings
  {
    std::string logFilePath;
    std::string logLevel;
    std::vector<std::pair<std::wstring, std::wstring>> pluginNamesAndPath;
    std::string pluginSearchPattern;
  };

  Settings& theCoreSettings();
  const toml::value* fetchPluginSettings(const wchar_t* pluginName);
}