#pragma once
#include <xlOil/DynamicRegister.h>

namespace xloil
{
  void registerLocalFuncs(
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcSpecs,
    const bool append);

  void clearLocalFunctions(const wchar_t* workbookName);
}