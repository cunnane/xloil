#pragma once
#include <xlOil/Register.h>
#include <xlOil/FuncSpec.h>
#include <memory>

namespace xloil
{
  class RegisteredWorksheetFunc
  {
  public:
    RegisteredWorksheetFunc(const std::shared_ptr<const WorksheetFuncSpec>& spec);

    ~RegisteredWorksheetFunc();

    /// <summary>
    /// De-registers the function in Excel and invalidates this object.
    /// Returns true on success.
    /// </summary>
    bool deregister();

    int registerId() const;

    const std::shared_ptr<const WorksheetFuncSpec>& spec() const;
    const std::shared_ptr<const FuncInfo>& info() const;

    /// <summary>
    /// Attempts some jiggery-pokery to avoid fully re-registering the function in Excel or 
    /// rebuilding the thunk code.  If it can't do this, de-registers the function and 
    /// returns false
    /// </summary>
    /// <returns>false if you need to call registerFunc</returns>
    virtual bool reregister(const std::shared_ptr<const WorksheetFuncSpec>& other);

  protected:
    int _registerId;
    std::shared_ptr<const WorksheetFuncSpec> _spec;
  };

  using RegisteredFuncPtr = std::shared_ptr<RegisteredWorksheetFunc>;

  RegisteredFuncPtr
    registerFunc(
      const std::shared_ptr<const WorksheetFuncSpec>& info) noexcept;

  int 
    registerFuncRaw(
      const std::shared_ptr<const FuncInfo>& info,
      const char* entryPoint,
      const wchar_t* moduleName);

  /// Remove a registered function. Zeros the passed pointer
  bool 
    deregisterFunc(const RegisteredFuncPtr& ptr);

  RegisteredFuncPtr
    findRegisteredFunc(const wchar_t* name);
}