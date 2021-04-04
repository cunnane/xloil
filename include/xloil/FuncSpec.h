#pragma once

#include "Register.h"
#include "ExportMacro.h"

namespace xloil
{
  class RegisteredWorksheetFunc;

  /// <summary>
  /// A base class which encapsulates the specification of a registered 
  /// function. That is, its <see cref="FuncInfo"/> and its call location.
  /// </summary>
  class WorksheetFuncSpec : public std::enable_shared_from_this<WorksheetFuncSpec>
  {
  public:
    WorksheetFuncSpec(const std::shared_ptr<const FuncInfo>& info) 
      : _info(info)
    {}

    /// <summary>
    /// Registers this function with the registry
    /// </summary>
    /// <returns>
    /// A <see cref="RegisteredWorksheetFunc"/> which can be used to deregister 
    /// this function.
    /// </returns>
    virtual std::shared_ptr<RegisteredWorksheetFunc> registerFunc() const = 0;

    /// <summary>
    /// Gets the <see cref="FuncInfo"/> associated with this WorksheetFuncSpec
    /// </summary>
    /// <returns></returns>
    const std::shared_ptr<const FuncInfo>& info() const { return _info; }

    /// <summary>
    /// Returns the name of the function, equivalent to info()->name
    /// </summary>
    /// <returns></returns>
    const std::wstring& name() const { return _info->name; }

  private:
    std::shared_ptr<const FuncInfo> _info;
  };

  // This class is used for statically registered functions and should
  // not be constructed directly.
  class StaticWorksheetFunction : public WorksheetFuncSpec
  {
  public:
    StaticWorksheetFunction(
      const std::shared_ptr<const FuncInfo>& info, 
      const std::wstring& dllName,
      const std::string& entryPoint)
      : WorksheetFuncSpec(info)
      , _dllName(dllName)
      , _entryPoint(entryPoint)
    {}

    XLOIL_EXPORT std::shared_ptr<RegisteredWorksheetFunc> registerFunc() const override;

    std::wstring _dllName;
    std::string _entryPoint;
  };
}