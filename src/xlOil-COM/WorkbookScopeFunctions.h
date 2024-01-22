#pragma once
#include <vector>
#include <string>
#include <memory>

namespace xloil
{
  struct FuncInfo;
  class LocalWorksheetFunc;
  
  constexpr wchar_t* theAutoGenModulePrefix = L"xlOil_";

  namespace COM
  {
    void writeLocalFunctionsToVBA(
      const wchar_t* workbookName,
      const std::vector<std::shared_ptr<const LocalWorksheetFunc>>& registeredFuncs,
      const wchar_t* vbaModuleName,
      const bool append);

    void removeExistingXlOilVBA(const wchar_t* workbookName);
  }
}