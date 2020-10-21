#pragma once
#include "Register.h"
#include "ExcelObj.h"

namespace xloil {
  class FuncSpec; 
  class RegisteredFunc; 
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

#define XLO_FUNC_END(func) \
    catch (const std::exception& err) \
    { \
      XLO_RETURN_ERROR(err); \
    } \
  } \
  XLO_REGISTER_FUNC(func)

#define XLO_RETURN_ERROR(err) return xloil::returnValue(err)

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

  struct FuncRegistrationMemo
  {
    typedef FuncRegistrationMemo self;
    FuncRegistrationMemo(
      const char* entryPoint_, size_t nArgs, const int* type);

    self& name(const wchar_t* txt)
    {
      _info->name = txt;
      return *this;
    }
    self& help(const wchar_t* txt)
    {
      _info->help = txt;
      return *this;
    }
    self& category(const wchar_t* txt)
    {
      _info->category = txt;
      return *this;
    }
    /// <summary>
    /// Specifies an arg as optional. This just effects the auto generated
    /// help string
    /// </summary>
    self& optArg(const wchar_t* name, const wchar_t* help = nullptr)
    {
      arg(name, help);
      _info->args[_iArg - 1].type |= FuncArg::Optional;
      return *this;
    }
    self& arg(const wchar_t* name, const wchar_t* help = nullptr)
    {
      if (_iArg >= _info->args.size())
        XLO_THROW("Too many args for function");
      auto& arg = _info->args[_iArg++];
      arg.name = name;
      if (help)
        arg.help = help;
      return *this;
    }
    self& command()
    {
      _info->options |= FuncInfo::COMMAND;
      return *this;
    }
    self& hidden()
    {
      _info->options |= FuncInfo::HIDDEN;
      return *this;
    }
    self& macro()
    {
      _info->options |= FuncInfo::MACRO_TYPE;
      return *this;
    }
    self& threadsafe()
    {
      _info->options |= FuncInfo::THREAD_SAFE;
      return *this;
    }
    // TODO: public but not exported...can we hide this?
    std::shared_ptr<FuncInfo> getInfo();

    std::string entryPoint;

  private:
    std::shared_ptr<FuncInfo> _info;
    size_t _iArg;
  };

  XLOIL_EXPORT FuncRegistrationMemo& 
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
    template<> struct ArgType<FPArray&> { static constexpr auto value = FuncArg::Array | FuncArg::ReturnVal; };
    template<> struct ArgType<const RangeArg&> { static constexpr auto value = FuncArg::Range; };

    template<class T> struct VoidArgType : public ArgType<T>
    {};
#ifdef XLOIL_UNSAFE_INPLACE_RETURN
    template<> struct VoidArgType<ExcelObj&> { static constexpr auto value = FuncArg::Obj | FuncArg::ReturnVal; };
    template<> struct VoidArgType<ExcelObj*> { static constexpr auto value = FuncArg::Obj | FuncArg::ReturnVal; };
#endif
    template<> struct VoidArgType<const AsyncHandle&> { static constexpr auto value = FuncArg::AsyncHandle; };

#ifndef _WIN64
#define XLOIL_STDCALL __stdcall
#else
#define XLOIL_STDCALL
#endif

    template<class T> struct ArgTypes;
    template <class Ret, class... Args>
    struct ArgTypes<Ret(XLOIL_STDCALL *)(Args...)>
    {
      static constexpr int types[sizeof...(Args)] = { ArgType<Args>::value ... };
      static constexpr size_t nArgs = sizeof...(Args);
    };
    template <class... Args>
    struct ArgTypes<void(XLOIL_STDCALL *)(Args...)>
    {
      static constexpr int types[sizeof...(Args)] = { VoidArgType<Args>::value ... };
      static constexpr size_t nArgs = sizeof...(Args);
    };
  }

  template <class TFunc> inline FuncRegistrationMemo&
    registrationMemo(const char* name, TFunc)
  {
    auto argTypes = detail::ArgTypes<TFunc>();
    return createRegistrationMemo(
      name, argTypes.nArgs, argTypes.types);
  }

  std::vector<std::shared_ptr<const FuncSpec>>
    processRegistryQueue(const wchar_t* moduleName);

  XLOIL_EXPORT std::vector<std::shared_ptr<const RegisteredFunc>>
    registerStaticFuncs(const wchar_t* moduleName, std::wstring& errors);
}