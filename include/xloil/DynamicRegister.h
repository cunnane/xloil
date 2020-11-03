#include <xloil/FuncSpec.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
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

    /// <summary>
    /// We want to skip the first argument as it will be of type FuncInfo.
    /// </summary>
    template <typename ReturnType, typename FirstArg, typename... Args>
    struct DropFirstArgTypes
    {
      static constexpr int types[sizeof...(Args)] =  { ArgType<Args>::value... };
      static constexpr size_t nArgs = sizeof...(Args);
    };

    template<class T> struct LambdaArgTypes
      : FunctionTraits<DropFirstArgTypes, T>
    {};
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

  /// <summary>
  /// Dynamically registers the provided callable, inheriting from 
  /// <see cref="FuncInfoBuilder"/> which allows customisation of the registration
  /// info.  Call <see cref="RegisterLambda::registerFunc"/> on this object to
  /// register the function.
  /// </summary>
  class RegisterLambda : public FuncInfoBuilder<RegisterLambda>
  {
    ExcelFuncObject _function;

    template<class TFunc, class TFirst, class TArray, size_t... Ints>
    static auto unpackArgs(
      TFunc func,
      TFirst firstArg,
      TArray args,
      std::index_sequence<Ints...>)
    {
      return func(firstArg, *args[Ints]...);
    }

  public:
    template <class TFunc>
    RegisterLambda(TFunc func)
      : FuncInfoBuilder(
          detail::LambdaArgTypes<TFunc>::nArgs,
          detail::LambdaArgTypes<TFunc>::types)
    {
      // TODO: check func is noexcept
      _function = [func](const FuncInfo& info, const ExcelObj** args)
      {
        return unpackArgs(
          func,
          info, 
          args,
          std::make_index_sequence<detail::LambdaArgTypes<TFunc>::nArgs>{});
      };
    }

    /// <summary>
    /// Registers this function and returns a handle to a <see cref="RegisteredFunc"/>
    /// object. Note that the handle must be kept in scope as its destructor
    /// unregisters the function.
    /// </summary>
    std::shared_ptr<RegisteredFunc> registerFunc()
    {
      // TODO: some things e.g. async not supported...check?
      return std::make_shared<LambdaFuncSpec>(getInfo(), _function)->registerFunc();
    }
  };
}