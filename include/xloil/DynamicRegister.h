#pragma once
#include <xloil/FuncSpec.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  template<typename TRet, typename TData> using DynamicCallback
    = TRet(*)(const TData* data, const ExcelObj**) noexcept;

  class DynamicSpec : public FuncSpec
  {
  public:
    template <class TRet, class TData>
    DynamicSpec(
      const std::shared_ptr<const FuncInfo>& info,
      DynamicCallback<TRet, TData> callback,
      const std::shared_ptr<const TData>& context)
      : DynamicSpec(
          info,
          (DynamicCallback<TRet, void>) callback,
          std::static_pointer_cast<const void>(context))
    {}

    DynamicSpec(
      const std::shared_ptr<const FuncInfo>& info,
      DynamicCallback<void, void> callback,
      const std::shared_ptr<const void>& context)
      : FuncSpec(info)
      , _callback(callback)
      , _context(context)
      , _hasReturn(false)
    {}

    DynamicSpec(
      const std::shared_ptr<const FuncInfo>& info,
      DynamicCallback<ExcelObj*, void> callback,
      const std::shared_ptr<const void>& context)
      : FuncSpec(info)
      , _callback(callback)
      , _context(context)
      , _hasReturn(true)
    {}

    XLOIL_EXPORT std::shared_ptr<RegisteredFunc> registerFunc() const override;

    //TODO: private:
    std::shared_ptr<const void> _context;
    void* _callback;
    bool _hasReturn;
  };


  namespace detail
  {
    /// <summary>
    /// We want to skip the first argument as it will be of type FuncInfo.
    /// </summary>
    template <typename ReturnType, typename FirstArg, typename... Args>
    struct DropFirstArgTypes
    {
      static constexpr int types[sizeof...(Args)] =  { ArgType<Args>::value... };
      static constexpr size_t nArgs = sizeof...(Args);
      
      template <size_t i> struct arg
      {
        using type = typename std::tuple_element<i, std::tuple<Args...>>::type;
      };
    };
    template <typename FirstArg, typename... Args>
    struct DropFirstArgTypes<void, FirstArg, Args...>
    {
      static constexpr int types[sizeof...(Args)] = { VoidArgType<Args>::value... };
      static constexpr size_t nArgs = sizeof...(Args);

      template <size_t i> struct arg
      {
        using type = typename std::tuple_element<i, std::tuple<Args...>>::type;
      };
    };
    template<class T> struct LambdaArgTypes
      : FunctionTraits<DropFirstArgTypes, T>
    {};

    template<typename TRet>
    struct DynamicCallbackFromLambda
    {
      template<typename TFunc, size_t... ArgIndices>
      auto operator()(TFunc func, std::index_sequence<ArgIndices...>)
      {
        return [func](const FuncInfo& info, const ExcelObj** args)
        {
          try
          {
            return func(info, (LambdaArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
          }
          catch (const std::exception& e)
          {
            return returnValue(e);
          }
        };
      }
    };

    template<>
    struct DynamicCallbackFromLambda<void>
    {
      template<typename TFunc, size_t... ArgIndices >
      auto operator()(TFunc func, std::index_sequence<ArgIndices...>)
      {
        return [func](const FuncInfo& info, const ExcelObj** args)
        {
          try
          {
            func(info, (LambdaArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
          }
          catch (...)
          {
          }
        };
      }
    };

    template<typename TRet, typename TFunc>
    auto dynamicCallbackFromLambda(TFunc func)
    {
      return DynamicCallbackFromLambda<TRet>()(
        func,
        std::make_index_sequence<LambdaArgTypes<TFunc>::nArgs>{});
    }
  }

  /// <summary>
  /// Constructs a FuncSpec from an std::function object which 
  /// takes <see cref="ExcelObj"/> arguments
  /// </summary>
  template<class TRet=ExcelObj*>
  class LambdaSpec : public FuncSpec
  {
  public:
    LambdaSpec(
      const std::shared_ptr<const FuncInfo>& info,
      const DynamicExcelFunc<TRet>& function)
      : FuncSpec(info)
      , function(function)
    {}

    XLOIL_EXPORT std::shared_ptr<RegisteredFunc> registerFunc() const override;

    auto call(const ExcelObj** args) const
    {
      return function(*info(), args);
    }
    DynamicExcelFunc<TRet> function;
  };

  /// <summary>
  /// Dynamically registers the provided callable, inheriting from 
  /// <see cref="FuncInfoBuilder"/> which allows customisation of the registration
  /// info.  Call <see cref="RegisterLambda::registerFunc"/> on this object to
  /// register the function.
  /// </summary>
  template<class TRet = ExcelObj*>
  class RegisterLambda : public FuncInfoBuilder<RegisterLambda<TRet>>
  {
    DynamicExcelFunc<TRet> _registerFunction;

  public:

    template <class TFunc>
    RegisterLambda(TFunc func)
      : FuncInfoBuilder(
          detail::LambdaArgTypes<TFunc>::nArgs,
          detail::LambdaArgTypes<TFunc>::types)
    {
      _registerFunction = detail::dynamicCallbackFromLambda<TRet, TFunc>(func);
    }

    /// <summary>
    /// Registers this function and returns a handle to a <see cref="RegisteredFunc"/>
    /// object. Note that the handle must be kept in scope as its destructor
    /// unregisters the function.
    /// </summary>
    std::shared_ptr<RegisteredFunc> registerFunc()
    {
      return std::make_shared<LambdaSpec<TRet>>(
        getInfo(), _registerFunction)->registerFunc();
    }
  };
}