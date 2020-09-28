#pragma once
#include <xlOil/Register.h>
#include <xlOil/FuncSpec.h>
#include <xlOil/Interface.h>
#include <memory>

namespace xloil
{
  class RegisteredFunc
  {
  public:
    RegisteredFunc(const std::shared_ptr<const FuncSpec>& spec);

    ~RegisteredFunc();

    /// <summary>
    /// De-registers the function in Excel and invalidates this object.
    /// Returns true on success.
    /// </summary>
    bool deregister();

    int registerId() const;

    const std::shared_ptr<const FuncSpec>& spec() const;
    const std::shared_ptr<const FuncInfo>& info() const;

    /// <summary>
    /// Attempts some jiggery-pokery to avoid fully re-registering the function in Excel or 
    /// rebuilding the thunk code.  If it can't do this, de-registers the function and 
    /// returns false
    /// </summary>
    /// <returns>false if you need to call registerFunc</returns>
    virtual bool reregister(const std::shared_ptr<const FuncSpec>& other);

  protected:
    int _registerId;
    std::shared_ptr<const FuncSpec> _spec;
  };

  using RegisteredFuncPtr = std::shared_ptr<RegisteredFunc>;

  RegisteredFuncPtr
    registerFunc(
      const std::shared_ptr<const FuncSpec>& info) noexcept;

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

  /// <summary>
  /// File source which collects and registers any declared
  /// static functions
  /// </summary>
  class StaticFunctionSource : public FileSource
  {
  public:
    StaticFunctionSource(const wchar_t* pluginPath);
  };
}