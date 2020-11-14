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
    template <typename ReturnType, typename... Args>
    struct ArgTypesDefs<ReturnType, const FuncInfo&, Args...> 
      : public ArgTypesDefs<ReturnType, Args...>
    {
      static constexpr bool hasInfo = true;
    };

    template<typename TFunc>
    auto hasInfo(int) -> decltype(ArgTypes<TFunc>::hasInfo, std::true_type());
    template<typename TFunc>
    auto hasInfo(long) -> std::false_type;

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
            if constexpr (decltype(hasInfo<TFunc>(0))::value)
              return func(info, (ArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
            else
              return func((ArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
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
            if constexpr (decltype(hasInfo<TFunc>(0))::value)
              func(info, (ArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
            else
              func((ArgTypes<TFunc>::arg<ArgIndices>::type)(*args[ArgIndices])...);
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
        std::make_index_sequence<ArgTypes<TFunc>::nArgs>{});
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
  class RegisterLambda : public FuncInfoBuilderT<RegisterLambda<TRet>>
  {
    DynamicExcelFunc<TRet> _registerFunction;

  public:
    /// <summary>
    /// Creates a lambda registration builder from a callable and optionally
    /// a FuncInfo. If a FuncInfo is not provided, it can be built up using the 
    /// FuncInfoBuilder methods on this class
    /// </summary>
    template <class TFunc>
    RegisterLambda(TFunc func, std::shared_ptr<FuncInfo> info = nullptr)
      : FuncInfoBuilderT(
          detail::ArgTypes<TFunc>::nArgs,
          detail::ArgTypes<TFunc>::types)
    {
      _registerFunction = detail::dynamicCallbackFromLambda<TRet, TFunc>(func);
      if (info)
      {
        if (info->numArgs() != _info->numArgs())
          XLO_THROW("RegisterLambda: if FuncInfo is provided, number of args must match");
        _info = info;
      }
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