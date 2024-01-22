#pragma once
#include <xloil/Interface.h>
#include <vector>
#include <string>

namespace xloil
{
  /// <summary>
  /// Load and attach the specified plugin to the given addin context.
  /// Called during autoOpen.
  /// </summary>
  /// <returns>True on sucess. A log entry will be written on failure</returns>
  bool loadPluginForAddin(AddinContext& context, const std::wstring& pluginName) noexcept;
  
  /// <summary>
  /// Detach the specified plugin from the addin context. In pratice this
  /// may not be called as Excel does not invoke autoClose on all XLLs
  /// during application exit.
  /// </summary>
  /// <returns>True on sucess. A log entry will be written on failure</returns>
  bool detachPluginForAddin(AddinContext& context, const std::wstring& pluginName) noexcept;

  /// Unloads any plugins prior to takedown of the Core XLL. 
  /// Called by xlAutoClose
  void unloadAllPlugins() noexcept;

  // Currently unused
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
