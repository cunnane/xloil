#pragma once
#include <xlOil/Register.h>
#include <memory>

namespace xloil
{
  class RegisteredFunc
  {
  public:
    /// <summary>
    /// Called internally
    /// </summary>
    RegisteredFunc(
      const std::shared_ptr<const FuncInfo>& info, 
      int registerId,
      const std::shared_ptr<void>& context, 
      void* thunk,
      size_t thunkSize);
    ~RegisteredFunc();

    /// <summary>
    /// De-registers the function in Excel and invalidates this object
    /// </summary>
    void deregister();

    int registerId() const;

    const std::shared_ptr<const FuncInfo>& info() const;

    /// <summary>
    /// Attempts some jiggery-pokery to avoid fully re-registering the function in Excel or 
    /// rebuilding the thunk code.  If it can't do this, de-registers the function and 
    /// returns false
    /// </summary>
    /// <param name="newInfo"></param>
    /// <param name="newContext"></param>
    /// <returns>false if you need to call registerFunc</returns>
    bool reregister(
      const std::shared_ptr<const FuncInfo>& newInfo, 
      const std::shared_ptr<void>& newContext);

  private:
    std::shared_ptr<const FuncInfo> _info;
    int _registerId;
    std::shared_ptr<void> _context;
    void* _thunk;
    size_t _thunkSize;
  };

  using RegisteredFuncPtr = std::shared_ptr<RegisteredFunc>;

  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, RegisterCallback callback, const std::shared_ptr<void>& context) noexcept;
  
  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, AsyncCallback callback, const std::shared_ptr<void>& context) noexcept;

  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, const char* functionName, const wchar_t* moduleName) noexcept;

  RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, const ExcelFuncPrototype& f) noexcept;


  template<class T> RegisteredFuncPtr
    registerFunc(const std::shared_ptr<const FuncInfo>& info, RegisterCallbackT<T> callback, const std::shared_ptr<T>& data) noexcept
  {
    return registerFunc(info, (RegisterCallback)callback, std::static_pointer_cast<void>(data));
  }

  /// Remove a registered function. Zeros the passed pointer
  void deregisterFunc(const std::shared_ptr<RegisteredFunc>& ptr);

  // TODO: the body of this is actually in Register.cpp
  std::vector<RegisteredFuncPtr> processRegistryQueue(const wchar_t* moduleName);
}