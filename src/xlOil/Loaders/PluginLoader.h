#pragma once
#include <xloil/Interface.h>
#include <vector>
#include <string>

namespace xloil
{
  class AddinContext;

  /// Load plugins according to settings file. Called by xlAutoOpen
  void loadPlugins(AddinContext* context, const std::vector<std::wstring>& names) noexcept;

  /// Unloads any plugins prior to takedown of the Core XLL. 
  /// Called by xlAutoClose
  void unloadAllPlugins() noexcept;

  std::vector<std::wstring> listPluginNames();


  /// <summary>
  /// File source which collects and registers any declared
  /// static functions
  /// </summary>
  class StaticFunctionSource : public FileSource
  {
  public:
    StaticFunctionSource(const wchar_t* pluginPath);
  };
}
