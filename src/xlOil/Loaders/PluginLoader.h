#pragma once
#include <xloil/Interface.h>
#include <vector>
#include <string>

namespace xloil
{
  /// Load plugins according to settings file. Called by xlAutoOpen
  void loadPluginsForAddin(AddinContext& context) noexcept;

  /// Unloads any plugins prior to takedown of the Core XLL. 
  /// Called by xlAutoClose
  void unloadAllPlugins() noexcept;

  std::vector<std::wstring> listPluginNames();


  /// <summary>
  /// File source which collects and registers any declared
  /// static functions
  /// </summary>
  class StaticFunctionSource : public FuncSource
  {
  public:
    /// <summary>
    /// This collects all statically declared Excel functions, i.e. raw C functions
    /// It assumes that this function and hence processRegistryQueue is run after each
    /// plugin has been loaded, so that all functions on the queue belong to the 
    /// current plugin
    /// </summary>
    /// 
    StaticFunctionSource(const wchar_t* pluginPath);

    const std::wstring& name() const override { return _sourcePath; }
    void init() override;
  private:
    std::wstring _sourcePath;
  };
}
