#include <xloil-xll/FuncRegistry.h>
#include <xloil/StaticRegister.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/Loaders/PluginLoader.h>

// See dodgy issues here: https://bugs.python.org/issue34309

namespace xloil
{
  bool reloadPlugin(const wchar_t* pluginName) noexcept
  {
    auto& loadedPlugins = getLoadedPlugins();
    auto found = loadedPlugins.find(pluginName);
    if (found == loadedPlugins.end())
      return false;

    XLO_INFO(L"Reloading plugin {}", pluginName);

    auto[name, plugin] = *found;

    rtdAsyncManagerClear();
    unloadPluginImpl(name.c_str(), plugin);
    loadedPlugins.erase(found);


    vector<wstring> names = { name };
    loadPlugins(plugin.Context, names);
    return true;
  }

  XLO_FUNC_START(xloReload(
    const ExcelObj& plugin
  ))
  {
    const auto pluginName = plugin.toString();
    return returnValue(reloadPlugin(pluginName.c_str()));
  }
  XLO_FUNC_END(xloReload).threadsafe()
    .help(L"Reloads a plugin")
    .arg(L"function", L"Name of plugin");
}