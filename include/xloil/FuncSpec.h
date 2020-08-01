#pragma once

#include "Register.h"
#include "ExportMacro.h"

namespace xloil
{
  class RegisteredFunc;

  /// <summary>
  /// A base class which encapsulates the specification of a registered 
  /// function. That is, its <see cref="FuncInfo"/> and its call location.
  /// </summary>
  class FuncSpec : public std::enable_shared_from_this<FuncSpec>
  {
  public:
    FuncSpec(const std::shared_ptr<const FuncInfo>& info) : _info(info) {}

    /// <summary>
    /// Registers this function with the registry
    /// </summary>
    /// <returns>
    /// A <see cref="RegisteredFunc"/> which can be used to deregister 
    /// this function.
    /// </returns>
    virtual std::shared_ptr<RegisteredFunc> registerFunc() const = 0;

    /// <summary>
    /// Gets the <see cref="FuncInfo"/> associated with this FuncSpec
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

    XLOIL_EXPORT std::shared_ptr<RegisteredFunc> registerFunc() const override;

    std::wstring _dllName;
    std::string _entryPoint;
  };

  namespace detail
  {
    template<class> struct callback_traits;
  }

  template <class TCallback>
  class GenericCallbackSpec : public FuncSpec
  {
  public:
    template <class TData>
    GenericCallbackSpec(
      const std::shared_ptr<const FuncInfo>& info,
      typename detail::callback_traits<TCallback>::template type<TData> callback,
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

    XLOIL_EXPORT std::shared_ptr<RegisteredFunc> registerFunc() const override;

  //TODO: private:
    std::shared_ptr<void> _context;
    TCallback _callback;
  };

  using CallbackSpec = GenericCallbackSpec<RegisterCallback>;
  using AsyncCallbackSpec = GenericCallbackSpec<AsyncCallback>;

  namespace detail
  {
    template<> struct callback_traits<RegisterCallback> 
    { 
      template<class T> using type = RegisterCallbackT<T>;
    };
    template<> struct callback_traits<AsyncCallback> 
    { 
      template<class T> using type = AsyncCallbackT<T>; 
    };
  }

  /// <summary>
  /// Constructs a FuncSpec from an std::function object which 
  /// takes <see cref="ExcelObj"/> arguments
  /// </summary>
  class LambdaFuncSpec : public FuncSpec
  {
  public:
    LambdaFuncSpec(
      const std::shared_ptr<const FuncInfo>& info,
      const ExcelFuncObject& function)
      : FuncSpec(info)
      , _function(function)
    {}

    XLOIL_EXPORT std::shared_ptr<RegisteredFunc> registerFunc() const override;

    ExcelObj* call(const ExcelObj** args) const
    {
      return _function(*info(), args);
    }

    ExcelFuncObject _function;
  };
}