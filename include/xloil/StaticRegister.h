#pragma once
#include <xlOil/Register.h>
#include <xlOil/ExcelObj.h>
#include <array>

namespace xloil {
  class WorksheetFuncSpec; 
  class RegisteredWorksheetFunc; 
  class FPArray; 
  class RangeArg; 
  struct AsyncHandle;
}

// In XLO_FUNC_START a separate declaration is needed to the function implementation
// to work around this quite serious MSVC compiler bug:
// https://stackoverflow.com/questions/45590594/generic-lambda-in-extern-c-function


/// <summary>
/// Marks the start of an function registered in Excel
/// </summary>
#define XLO_FUNC_START(func) \
  XLO_ENTRY_POINT(XLOIL_XLOPER*) func; \
  XLOIL_XLOPER* __stdcall func \
  { \
    try 

#ifdef XLO_RETURN_COM_ERROR
#define XLO_FUNC_END(func) \
    XLO_RETURN_COM_ERROR \
    catch (const std::exception& err) \
    { \
      return xloil::returnValue(err); \
    } \
    catch (...) \
    { \
      return xloil::returnValue(xloil::CellError::Value); \
    } \
  } \
  XLO_REGISTER_FUNC(func)
#else
#define XLO_FUNC_END(func) \
    catch (const std::exception& err) \
    { \
      return xloil::returnValue(err); \
    } \
    catch (...) \
    { \
      return xloil::returnValue(xloil::CellError::Value); \
    } \
  } \
  XLO_REGISTER_FUNC(func)
#endif // XLO_RETURN_COM_ERROR

#define XLO_RETURN_ERROR(err) 

#define XLO_REGISTER_FUNC(func) extern auto _xlo_register_##func = xloil::registrationMemo(#func, func)

namespace xloil
{
  /// <summary>
   /// Constructs an ExcelObj from the given arguments, setting a flag to tell 
   /// Excel that xlOil will need a callback to free the memory. **This method must
   /// be used for final object passed back to Excel. It must not be used anywhere
   /// else**.
   /// </summary>
  template<class... Args>
  inline ExcelObj* returnValue(Args&&... args)
  {
    return (new ExcelObj(std::forward<Args>(args)...))->toExcel();
  }
  inline ExcelObj* returnValue(CellError err)
  {
    return const_cast<ExcelObj*>(&Const::Error(err));
  }
  inline ExcelObj* returnValue(const std::exception& e)
  {
    return returnValue(e.what());
  }
  inline ExcelObj* returnReference(const ExcelObj& obj)
  {
    return const_cast<ExcelObj*>(&obj);
  }
  inline ExcelObj* returnReference(ExcelObj& obj)
  {
    return &obj;
  }
  namespace detail
  {
    struct FuncInfoBuilderBase
    {
      FuncInfoBuilderBase(size_t nArgs, const int* types);
      template<int N> 
      FuncInfoBuilderBase(const std::array<int, N>& types)
        : FuncInfoBuilderBase(N, types.data())
      {}

      std::shared_ptr<FuncInfo> getInfo();

    protected:
      std::shared_ptr<FuncInfo> _info;
      size_t _iArg;
    };
  }

  template<class TSuper>
  struct FuncInfoBuilderT : public detail::FuncInfoBuilderBase
  {
    using self = TSuper;
    using detail::FuncInfoBuilderBase::FuncInfoBuilderBase;

    /// <summary>
    /// Sets the name for the function registration
    /// </summary>
    /// <param name="txt"></param>
    template<class T> self& name(T txt)
    {
      _info->name = txt;
      return cast();
    }
    /// <summary>
    /// Sets the function wizard help (max 255 chars)
    /// </summary>
    /// <param name="txt"></param>
    template<class T> self& help(T txt)
    {
      _info->help = txt;
      return cast();
    }
    /// <summary>
    /// Sets the category the function will be placed under in the wizard
    /// </summary>
    /// <param name="txt"></param>
    template<class T> self& category(T txt)
    {
      _info->category = txt;
      return cast();
    }
    /// <summary>
    /// Sets the name and optionally help for the next parameter and
    /// marks the arg as optional. This just effects the auto generated
    /// help string
    /// </summary>
    self& optArg(const wchar_t* name, const wchar_t* help = nullptr)
    {
      return optArg(_iArg++, name, help);
    }
    /// <summary>
    /// Sets the name and optionally help for the next parameter, incrementing
    /// the internal counter
    /// </summary>
    /// <param name="name"></param>
    /// <param name="help"></param>
    self& arg(const wchar_t* name, const wchar_t* help = nullptr)
    {
      return arg(_iArg++, name, help);
    }
    /// <summary>
    /// Sets the name and optionally help for the i-th parameter
    /// </summary>
    /// <param name="i"></param>
    /// <param name="name"></param>
    /// <param name="help">optional, may be null</param>
    self& arg(size_t i, const wchar_t* name, const wchar_t* help = nullptr)
    {
      if (i >= _info->args.size())
        XLO_THROW("Too many args for function");
      auto& arg = _info->args[i];
      arg.name = name;
      if (help)
        arg.help = help;
      return cast();
    }
    /// <summary>
    /// Sets the name and optionally help for the i-th parameter and
    /// marks the arg as optional. This just effects the auto generated
    /// help string
    /// </summary>
    /// <param name="i"></param>
    /// <param name="name"></param>
    /// <param name="help">optional, may be null</param>
    self& optArg(size_t i, const wchar_t* name, const wchar_t* help = nullptr)
    {
      arg(i, name, help);
      _info->args[i].type |= FuncArg::Optional;
      return cast();
    }

    self& command()
    {
      _info->options |= FuncInfo::COMMAND;
      return cast();
    }
    /// <summary>
    /// Hides the function in the function wizard
    /// </summary>
    self& hidden()
    {
      _info->options |= FuncInfo::HIDDEN;
      return cast();
    }
    /// <summary>
    /// Gives the function macro-sheet access rights to the API
    /// </summary>
    self& macro()
    {
      _info->options |= FuncInfo::MACRO_TYPE;
      return cast();
    }
    /// <summary>
    /// Tells Excel that the function is re-entrant and can safely be 
    /// called on multiple threads simultaneously
    /// </summary>
    self& threadsafe()
    {
      _info->options |= FuncInfo::THREAD_SAFE;
      return cast();
    }

  protected:
    self& cast() { return static_cast<self&>(*this); }
  };

  struct FuncInfoBuilder : public FuncInfoBuilderT<FuncInfoBuilder> {};

  struct StaticRegistrationBuilder : public FuncInfoBuilderT<StaticRegistrationBuilder>
  {
    StaticRegistrationBuilder(
      const char* entryPoint_, size_t nArgs, const int* type)
      : FuncInfoBuilderT(nArgs, type)
    {
      entryPoint = entryPoint_;
      name(utf8ToUtf16(entryPoint_));
    }

    std::string entryPoint;
  };

  XLOIL_EXPORT StaticRegistrationBuilder& 
    createRegistrationMemo(
      const char* entryPoint_, size_t nArgs, const int* types);

#if DOXYGEN
/// <summary>
/// Returning ExcelObj in-place is disabled by default. In the words of the XLL SDK:
/// 
/// "Excel permits the registration of functions that return an XLOPER by modifying 
/// an argument in place. However, if an XLOPER argument points to memory, and the 
/// pointer is then overwritten by the return value of the DLL function, Excel can 
/// leak memory. If the DLL allocated memory for the return value, Excel might try 
/// to free that memory, which could cause an immediate crash.  Therefore, you should 
/// not modify XLOPER/XLOPER12 arguments in place."
/// 
/// In practice, it can be safe to modify an ExcelObj in place, for instance xloSort
/// does this by changing the row order in the array, but without changing memory 
/// allocation.
/// </summary>
#define XLOIL_UNSAFE_INPLACE_RETURN
#endif

  namespace detail
  {
    template<class T> struct ArgType {};
    template<> struct ArgType<const ExcelObj&> { static constexpr auto value = FuncArg::Obj; };
    template<> struct ArgType<const ExcelObj*> { static constexpr auto value = FuncArg::Obj; };
    template<> struct ArgType<const FPArray&> { static constexpr auto value = FuncArg::Array; };
    template<> struct ArgType<const RangeArg&> { static constexpr auto value = FuncArg::Range; };

    /// <summary>
    /// In-place and async return argument types are only valid when the function 
    /// returns void
    /// </summary>
    template<class T> struct VoidArgType : public ArgType<T>
    {};
#ifdef XLOIL_UNSAFE_INPLACE_RETURN
    template<> struct VoidArgType<ExcelObj&> { static constexpr auto value = FuncArg::Obj | FuncArg::ReturnVal; };
    template<> struct VoidArgType<ExcelObj*> { static constexpr auto value = FuncArg::Obj | FuncArg::ReturnVal; };
#endif
    template<> struct VoidArgType<FPArray&> { static constexpr auto value = FuncArg::Array | FuncArg::ReturnVal; };
    template<> struct VoidArgType<const AsyncHandle&> { static constexpr auto value = FuncArg::AsyncHandle; };

#ifndef _WIN64
#define XLOIL_STDCALL __stdcall
#else
#define XLOIL_STDCALL
#endif

    /// <summary>
    /// Ultimately inherits from Defs<ReturnType, Args...> but due to the myriad
    /// ways which a callable can be expressed in C++, has a lot of specialisations
    /// </summary>
    template <template<typename, typename...> typename Defs, typename T>
    struct FunctionTraitsFilter;

    template <template<typename, typename...> typename Defs,
      typename ReturnType, typename... Args>
      struct FunctionTraitsFilter<Defs, ReturnType(Args...)>
      : Defs<ReturnType, Args...> {};

    template <template<typename, typename...> typename Defs,
      typename ReturnType, typename... Args>
      struct FunctionTraitsFilter<Defs, ReturnType(XLOIL_STDCALL *)(Args...)>
      : Defs<ReturnType, Args...> {};

    template <template<typename, typename...> typename Defs,
      typename ReturnType, typename ClassType, typename... Args>
      struct FunctionTraitsFilter<Defs, ReturnType(ClassType::*)(Args...)>
      : Defs<ReturnType, Args...> {};

    template <template<typename, typename...> typename Defs,
      typename ReturnType, typename ClassType, typename... Args>
      struct FunctionTraitsFilter<Defs, ReturnType(ClassType::*)(Args...) const>
      : Defs<ReturnType, Args...> {};

    template <template<typename, typename...> typename Defs, typename T, typename SFINAE = void>
    struct FunctionTraits
      : FunctionTraitsFilter<Defs, T> {};

    template <template<typename, typename...> typename Defs, typename T>
    struct FunctionTraits<Defs, T, decltype((void)&T::operator())>
      : FunctionTraitsFilter<Defs, decltype(&T::operator())> {};

    template <typename ReturnType, typename... Args>
    struct ArgTypesDefs
    {
      static constexpr std::array<int, sizeof...(Args)> types = { ArgType<Args>::value ... };
      static constexpr size_t nArgs = sizeof...(Args);
      template <size_t i> struct arg
      {
        using type = typename std::tuple_element<i, std::tuple<Args...>>::type;
      };
    };
    template <typename... Args>
    struct ArgTypesDefs<void, Args...>
    {
      static constexpr std::array<int, sizeof...(Args)> types = { VoidArgType<Args>::value ... };
      static constexpr size_t nArgs = sizeof...(Args);
      template <size_t i> struct arg
      {
        using type = typename std::tuple_element<i, std::tuple<Args...>>::type;
      };
    };

    template<class T> struct ArgTypes
      : FunctionTraits<ArgTypesDefs, T>
    {};
  }

  template <class TFunc> inline StaticRegistrationBuilder&
    registrationMemo(const char* name, TFunc)
  {
    auto argTypes = detail::ArgTypes<TFunc>();
    return createRegistrationMemo(
      name, argTypes.nArgs, argTypes.types.data());
  }

  std::vector<std::shared_ptr<const WorksheetFuncSpec>>
    processRegistryQueue(const wchar_t* moduleName);

  XLOIL_EXPORT std::vector<std::shared_ptr<const RegisteredWorksheetFunc>>
    registerStaticFuncs(const wchar_t* moduleName, std::wstring& errors);
}