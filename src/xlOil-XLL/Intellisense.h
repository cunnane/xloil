#pragma once
#include <memory>
#include <vector>

namespace xloil
{
  struct FuncInfo;
  /// <summary>
  /// Will fail, possibly spectacularly, unless called in XLL context
  /// </summary>
  void publishIntellisenseInfo(const std::shared_ptr<const FuncInfo>& info);
  void publishIntellisenseInfo(const std::vector<std::shared_ptr<const FuncInfo>>& infos);
  void registerIntellisenseHook(const wchar_t* xllPath);
}