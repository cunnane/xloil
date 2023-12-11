#include "Settings.h"
#include "Exception.h"
#include <xlOil/StringUtils.h>
#include <xloilHelpers/Environment.h>

#include <filesystem>
#include <fstream>

// We need access to a private ctor - yikes!
#define private public
#include <toml++/toml.h>
#undef private

namespace fs = std::filesystem;
using xloil::Helpers::Exception;
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
      auto findStr(const toml::view_node& root, const char* tag, const string& default)
      {
        return root[tag].value_or(default);
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
    std::wstring logFilePath(const toml::table& root)
    {
      auto found = findStr(root[XLOIL_SETTINGS_ADDIN_SECTION], "LogFile", "");
      return !found.empty()
        ? utf8ToUtf16(found)
        : fs::path(utf8ToUtf16(*root.source().path)).replace_extension("log").wstring();
    }
    std::string logLevel(const toml::view_node& root)
    {
      return findStr(root, "LogLevel", "warn");
    }
    std::string logPopupLevel(const toml::view_node& root)
    {
      return findStr(root, "LogPopupLevel", "error");
    }
    std::pair<size_t, size_t> Settings::logRotation(const toml::view_node& root)
    {
      // (size_t) cast needed for 32-bit as TOML lib is hard-coded to 
      // return int64 for all integer types
      return std::make_pair(
        (size_t)root["LogMaxSize"].value_or(1024u),
        (size_t)root["LogNumberOfFiles"].value_or(2u));
    }
    std::vector<std::wstring> dateFormats(const toml::view_node& root)
    {
      return findVecStr(root, "DateFormats");
    }

    namespace
    {
      // Settings in the enviroment block looks like key=val
      // We interpret this as an environment variable to set
      template<class TTable, class TResult>
      void writeTableAsPairs(const TTable& table, TResult& container)
      {
        for (auto [key, val] : table)
        {
          container.emplace_back(make_pair(
            utf8ToUtf16(key),
            utf8ToUtf16(val.value_or(""))));
        }
      }
    }

    std::vector<std::pair<std::wstring, std::wstring>> 
      environmentVariables(const toml::view_node& root)
    {
      vector<pair<wstring, wstring>> result;
      auto environment = root["Environment"];
      if (environment.is_array())
      {
        for (auto& table : *environment.as_array())
          if (table.is_table())
            writeTableAsPairs(*table.as_table(), result);
      }
      else if (environment.is_table())
      {
        writeTableAsPairs(*environment.as_table(), result);
      }
      return result;
    }

    bool loadBeforeCore(const toml::table& root)
    {
      return root[XLOIL_SETTINGS_ADDIN_SECTION]["LoadBeforeCore"].value_or(false);
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
          if (_stricmp((*i).first.data(), name) == 0)
            return toml::view_node(&(*i).second);
        }
   
      return toml::view_node();
    }
  }
  std::shared_ptr<const toml::table> findSettingsFile(const wchar_t* dllPath)
  {
    fs::path path;
 
    const auto settingsFileName = 
      fs::path(dllPath).filename().replace_extension(XLOIL_SETTINGS_FILE_EXT);
    
    auto directoryOverride = getEnvironmentVar("XLOIL_SETTINGS_DIR");
    if (!directoryOverride.empty())
      path = fs::path(directoryOverride) / settingsFileName;
    else
    {
      // Look in the user's appdata
      path = fs::path(getEnvironmentVar(L"APPDATA")) / L"xlOil" / settingsFileName;
    }
    
    // If not found, check the same directory as the dll itself
    std::error_code fsErr;
    if (!fs::exists(path, fsErr))
      path = fs::path(dllPath).remove_filename() / settingsFileName;

    try
    {
      if (!fs::exists(path, fsErr))
        return shared_ptr<const toml::table>();
      
      auto ifs = std::ifstream{ path.wstring() };

      return make_shared<toml::table>(
        toml::parse(ifs, utf16ToUtf8(path.wstring())));
    }
    catch (const toml::parse_error& e)
    {
      throw Exception("Error parsing settings file '%s' at line %d:\n %s",
        path.string().c_str(), e.source().begin.line, e.what());
    }
  }
}