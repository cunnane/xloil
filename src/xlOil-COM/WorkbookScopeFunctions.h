#pragma once
#include <vector>
#include <string>

namespace xloil
{
  class WorksheetFuncSpec;

  namespace COM
  {
    void writeLocalFunctionsToVBA(
      const wchar_t* workbookName,
      const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& registeredFuncs);
  }
}