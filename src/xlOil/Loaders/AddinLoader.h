#pragma once
#include <memory>

namespace xloil
{
  class AddinContext; class FileSource;

  AddinContext* theCoreContext();

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    findFileSource(const wchar_t* sourcePath);

  void
    deleteFileSource(const std::shared_ptr<FileSource>& source);

  bool openXll(const wchar_t* xllPath);
  void closeXll(const wchar_t* xllPath);
}