#pragma once
#include <vector>
#include <string>
#include <memory>

namespace toml {
    class table;
    class node;
    template <typename> class node_view;
    using view_node = node_view<const node>;
}

namespace xloil
{
  constexpr char* XLOIL_SETTINGS_FILE_EXT = "ini";
  constexpr char* XLOIL_SETTINGS_ADDIN_SECTION = "Addin";

  namespace Settings
  {
    std::wstring logFilePath(const toml::table& root);

    std::string logLevel(const toml::table& root);

    std::string logPopupLevel(const toml::table& root);

    std::string logFlushLevel(const toml::table& root);

    std::pair<size_t, size_t> logRotation(const toml::table& root);

    std::vector<std::wstring> plugins(const toml::view_node& root);

    std::wstring pluginSearchPattern(const toml::view_node& root);

    std::vector<std::wstring> dateFormats(const toml::table& root);

    std::vector<std::pair<std::wstring, std::wstring>>
      environmentVariables(const toml::view_node& root);

    bool loadBeforeCore(const toml::table& root);

    /// <summary>
    /// Lookup name in table in a case-insensitive way. TOML lookup is case 
    /// sensitive because the creator "prefers it that way". That's fine, but 
    /// Microsoft thinks differently and so since 'name' is a filename, case
    /// sensitive lookup would be fairly astonishing.
    /// </summary>
    toml::node_view<const toml::node> findPluginSettings(
      const toml::table* table, const char* name);
  };

  std::shared_ptr<const toml::table>
    findSettingsFile(const wchar_t* dllPath);
}