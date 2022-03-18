#pragma once
#include <memory>

namespace xloil
{
  class AddinContext; class FileSource;

  AddinContext& theCoreContext();

  /// <summary>
  /// Called by the core entry point to initialise all the xloil.dll
  /// paths and register functions
  /// </summary>
  void createCoreContext();

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    findFileSource(const wchar_t* sourcePath);

  /// <summary>
  /// Removes the file source from all add-in contexts
  /// </summary>
  void deleteFileSource(const std::shared_ptr<FileSource>& source);

  /// <summary>
  /// Called by the core entry point to initialise each XLL except xloil.xll
  /// The core DLL is initialised by createCoreContext, not this function
  /// </summary>
  AddinContext& addinOpenXll(const wchar_t* xllPath);

  void loadPluginsForAddin(AddinContext& ctx);

  /// <summary>
  /// Triggered by xlAutoClose for each addin. When the last XLL is closed
  /// a teardown is initiated.
  /// </summary>
  void addinCloseXll(const wchar_t* xllPath);
}