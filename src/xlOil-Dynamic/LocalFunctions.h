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

  /// <summary>
  /// Action when registering local functions:
  ///   * Append to name module if possible. Append can only take place when 
  ///     the new functions do not share names with the existing ones
  ///   * Replace with a new VBA module, de-registering all existing local 
  ///     functions
  ///   * Clear all xlOil local function stubs in this workbook (prefixed with
  ///     xlOil_) before registering functions (implies Replace)
  /// </summary>
  enum class LocalFuncs
  {
    APPEND_MODULE,
    REPLACE_MODULE,
    CLEAR_MODULES
  };

  void registerLocalFuncs(
    std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>& existing,
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcs,
    const wchar_t* vbaModuleName,
    const LocalFuncs action);

  void clearLocalFunctions(
    std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>& existing);
  
  /// <summary>
  /// Returns true if a local function is currently executing
  /// </summary>
  bool isExecutingLocalFunction();
}