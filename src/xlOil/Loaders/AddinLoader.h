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
  std::shared_ptr<AddinContext> createCoreAddinContext();

  const std::map<std::wstring, std::shared_ptr<AddinContext>>& currentAddinContexts();

  /// <summary>
  /// Called by the core entry point to initialise each XLL except xloil.xll
  /// The core DLL is initialised by createCoreContext, not this function
  /// </summary>
  std::shared_ptr<AddinContext> createAddinContext(const wchar_t* xllPath);

  /// <summary>
  /// Triggered by xlAutoClose for each addin. When the last XLL is closed
  /// a teardown is initiated.
  /// </summary>
  void addinCloseXll(const wchar_t* xllPath);

  /// <summary>
  /// When Excel closes it may not call autoClose for Excel XLL. This function
  /// ensures tidy-up happens at DLL detach in a safe way which does not try
  /// to use the Excel APIs.
  /// </summary>
  void teardownAddinContext();
}