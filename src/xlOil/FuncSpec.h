#pragma once

#include "Register.h"
#include "ExportMacro.h"
namespace xloil
{
  class RegisteredFunc;

  class FuncSpec : public std::enable_shared_from_this<FuncSpec>
  {
  public:
    FuncSpec(const std::shared_ptr<const FuncInfo>& info) : _info(info) {}
    virtual std::shared_ptr<RegisteredFunc> registerFunc() const = 0;
    const std::shared_ptr<const FuncInfo>& info() const { return _info; }
    const std::wstring& name() const { return _info->name; }
  private:
    std::shared_ptr<const FuncInfo> _info;
  };

  class StaticSpec : public FuncSpec
  {
  public:
    StaticSpec(
      const std::shared_ptr<const FuncInfo>& info, 
      const std::wstring& dllName,
      const std::string& entryPoint)
      : FuncSpec(info)
      , _dllName(dllName)
      , _entryPoint(entryPoint)
    {}

    XLOIL_EXPORT virtual std::shared_ptr<RegisteredFunc> registerFunc() const;

    std::wstring _dllName;
    std::string _entryPoint;
  };
 
  template<class> struct callback_traits;

  template <class TCallback>
  class GenericCallbackSpec : public FuncSpec
  {
  public:
    template <class TData>
    GenericCallbackSpec(
      const std::shared_ptr<const FuncInfo>& info,
      typename callback_traits<TCallback>::template type<TData> callback,
      std::shared_ptr<TData> context)
      : GenericCallbackSpec(
        info, 
        (TCallback)callback, 
        std::static_pointer_cast<void>(context))
    {}

    GenericCallbackSpec(
      const std::shared_ptr<const FuncInfo>& info,
      TCallback callback,
      std::shared_ptr<void> context)
      : FuncSpec(info)
      , _callback(callback)
      , _context(context)
    {}

    XLOIL_EXPORT virtual std::shared_ptr<RegisteredFunc> registerFunc() const;

  //TODO: private:
    std::shared_ptr<void> _context;
    TCallback _callback;
  };

  using CallbackSpec = GenericCallbackSpec<RegisterCallback>;
  using AsyncCallbackSpec = GenericCallbackSpec<AsyncCallback>;

  template<> struct callback_traits<RegisterCallback> { template<class T> using type = RegisterCallbackT<T>; };
  template<> struct callback_traits<AsyncCallback> { template<class T> using type = AsyncCallbackT<T>; };

  class FuncObjSpec : public FuncSpec
  {
  public:
    FuncObjSpec(
      const std::shared_ptr<const FuncInfo>& info,
      const ExcelFuncObject& function)
      : FuncSpec(info)
      , _function(function)
    {}

    XLOIL_EXPORT virtual std::shared_ptr<RegisteredFunc> registerFunc() const;

    ExcelFuncObject _function;
  };
}