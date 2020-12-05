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
    /// <summary>
    /// This collects all statically declared Excel functions, i.e. raw C functions
    /// It assumes that this function and hence processRegistryQueue is run after each
    /// plugin has been loaded, so that all functions on the queue belong to the 
    /// current plugin
    /// </summary>
    void registerQueue();
  };
}
