#pragma once
#include <xlOil/Register.h>
#include <xlOil/FuncSpec.h>
#include <memory>
#include <map>

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

    /// <summary>
    /// 'Forgets' the registration - this stops any attempt to deregister the function
    /// when the object is destroyed.
    /// </summary>
    void forget() { _registerId = 0; }

  protected:
    int _registerId;
    std::shared_ptr<const WorksheetFuncSpec> _spec;
  };

  using RegisteredFuncPtr = std::shared_ptr<RegisteredWorksheetFunc>;

  /// <summary>
  /// Will fail unless called in XLL context
  /// </summary>
  /// <param name="info"></param>
  /// <returns></returns>
  RegisteredFuncPtr
    registerFunc(
      const std::shared_ptr<const WorksheetFuncSpec>& info) noexcept;

  int 
    registerFuncRaw(
      const std::shared_ptr<const FuncInfo>& info,
      const char* entryPoint,
      const wchar_t* moduleName);

  const std::map<std::wstring, RegisteredFuncPtr>& 
    registeredFuncsByName();

  /// <summary>
  /// Called during teardown to release resources used by registered
  /// functions. Does not attempt the deregister the functions as 
  /// the Excel APIs are not available after autoClose.
  /// </summary>
  void teardownFunctionRegistry();
}