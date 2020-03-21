#pragma once
#include <vector>
#include <string>

namespace xloil
{
  struct FuncInfo;

  void writeLocalFunctionsToVBA(
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const FuncInfo>>& registeredFuncs,
    const std::vector<std::wstring> coreRedirects);
}