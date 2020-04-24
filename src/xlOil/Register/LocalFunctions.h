#pragma once
#include <xlOil/Register.h>

namespace xloil
{
  void registerLocalFuncs(
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const FuncInfo>>& registeredFuncs,
    const std::vector<ExcelFuncObject> funcs);

  void forgetLocalFunctions(const wchar_t* workbookName);
}