#pragma once
#include <memory>

namespace xloil
{
  class AddinContext; class FileSource;

  AddinContext* theCoreContext();
  void createCoreContext();

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    findFileSource(const wchar_t* sourcePath);

  /// <summary>
  /// Removes the file source from all add-in contexts
  /// </summary>
  void deleteFileSource(const std::shared_ptr<FileSource>& source);

  AddinContext* openXll(const wchar_t* xllPath);
  void loadPluginsForAddin(AddinContext* ctx);
  void closeXll(const wchar_t* xllPath);
}