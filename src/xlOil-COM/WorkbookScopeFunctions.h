#pragma once
#include <vector>
#include <string>
#include <memory>

namespace xloil
{
  struct FuncInfo;

  namespace COM
  {
    void writeLocalFunctionsToVBA(
      const wchar_t* workbookName,
      const std::vector<std::shared_ptr<const FuncInfo>>& registeredFuncs,
      const bool append);
  }
}