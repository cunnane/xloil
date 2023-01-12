#pragma once
#include <memory>
#include <string>
#include <map>

namespace xloil
{
  class AddinContext; class FileSource; class FuncSource;

  AddinContext& theCoreContext();

  /// <summary>
  /// Called by the core entry point to initialise all the xloil.dll
  /// paths and register functions
  /// </summary>
  AddinContext& createCoreAddinContext(
    const std::shared_ptr<FuncSource>& staticFunctions);

  const std::map<std::wstring, std::shared_ptr<AddinContext>>& currentAddinContexts();

  /// <summary>
  /// Called by the core entry point to initialise each XLL except xloil.xll
  /// The core DLL is initialised by createCoreContext, not this function
  /// </summary>
  AddinContext& createAddinContext(const wchar_t* xllPath);

  /// <summary>
  /// Triggered by xlAutoClose for each addin. When the last XLL is closed
  /// a teardown is initiated.
  /// </summary>
  void addinCloseXll(const wchar_t* xllPath);
}