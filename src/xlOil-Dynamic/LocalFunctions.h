#pragma once
#include <xlOil/DynamicRegister.h>
#include <map>

namespace xloil
{
  class LocalWorksheetFunc
  {
  public:
    LocalWorksheetFunc(const std::shared_ptr<const WorksheetFuncSpec>& spec);

    ~LocalWorksheetFunc();

    intptr_t registerId() const;

    const std::shared_ptr<const WorksheetFuncSpec>& spec() const;
    const std::shared_ptr<const FuncInfo>& info() const;

  private:
    std::shared_ptr<const WorksheetFuncSpec> _spec;
  };

  void registerLocalFuncs(
    std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>& existing,
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcs,
    const bool append);

  void clearLocalFunctions(
    std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>& existing);
}