#pragma once
#include "Options.h"
#include "ExportMacro.h"
#include "Log.h"
#include <vector>
#include <memory>
#include <list>

#define XLO_FUNC XLO_ENTRY_POINT(XLOIL_XLOPER*)
#define XLO_REGISTER(func) extern auto _xlo_register_##func = xloil::registrationMemo(#func, func)

namespace xloil { class ExcelObj; }

namespace xloil
{
  class ExcelObj;

  /// <summary>
  /// Holds the description of an Excel function argument
  /// </summary>
  struct FuncArg
  {
    FuncArg(const wchar_t* name_, const wchar_t* help_ = nullptr)
      : name(name_ ? name_ : L"")
      , help(help_ ? help_ : L"")
    {}
    /// <summary>
    /// The name of the argument shown in the function wizard.
    /// </summary>
    std::wstring name;
    /// <summary>
    /// An optional help string for the argument displayed in the function wizard.
    /// </summary>
    std::wstring help;

    bool operator==(const FuncArg& that) const
    {
      return name == that.name && help == that.help;
    }
  };

  struct FuncInfo
  {
    /// <summary>
    /// 
    /// </summary>
    enum FuncOpts
    {
      THREAD_SAFE = 1 << 0,
      MACRO_TYPE  = 1 << 1,
      VOLATILE    = 1 << 2,
      COMMAND     = 1 << 3,
      ASYNC       = 1 << 4
    };

    XLOIL_EXPORT virtual ~FuncInfo();
    XLOIL_EXPORT bool operator==(const FuncInfo& that) const;
    bool operator!=(const FuncInfo& that) const { return !(*this == that); }

    /// <summary>
    /// The name of the function which will be used in worksheet (and VBA) calls.
    /// </summary>
    std::wstring name;

    /// <summary>
    /// An optional help string which appears in the function wizard.
    /// </summary>
    std::wstring help;

    /// <summary>
    /// An optional category group which determiines where the function appears in
    /// the funcion wizard.
    /// </summary>
    std::wstring category;
    
    /// <summary>
    /// Array of function arguments in the order you expect them to be passed
    /// Excel does not support named keyword arguments, but a <see cref="ExcelDict"/>
    /// can be used to give this behaviour. Can be empty.
    /// </summary>
    std::vector<FuncArg> args;

    // TODO: make me an Enum, apparently that's OK in modern c++
    int options;
    size_t numArgs() const { return args.size(); }
  };

  template<class TData> using RegisterCallbackT = ExcelObj* (*)(TData* data, const ExcelObj**);
  typedef RegisterCallbackT<void> RegisterCallback;

  template<class TData> using AsyncCallbackT = void (*)(TData* data, const ExcelObj*, const ExcelObj**);
  typedef AsyncCallbackT<void> AsyncCallback;

  using ExcelFuncPrototype = std::function<ExcelObj*(const FuncInfo& info, const ExcelObj**)>;

  struct FuncRegistrationMemo
  {
    typedef FuncRegistrationMemo self;
    FuncRegistrationMemo(const char* entryPoint_, size_t nArgs);

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
    self& arg(const wchar_t* name, const wchar_t* help=nullptr)
    {
      _info->args.emplace_back(FuncArg(name, help));
      return *this;
    }
    self& async()
    {
      _info->options |= FuncInfo::ASYNC;
      return *this;
    }
    self& command()
    {
      _info->options |= FuncInfo::COMMAND;
      return *this;
    }
    self& macro()
    {
      _info->options |= FuncInfo::MACRO_TYPE;
      return *this;
    }
    // TODO: public but not exported...can we hide this?
    std::shared_ptr<const FuncInfo> getInfo();

    std::string entryPoint;

  private:
    std::shared_ptr<FuncInfo> _info;
    size_t _nArgs;
  };

  template <class R, class... Args> constexpr size_t
    getArgumentCount(R(*)(Args...))
  {
    return sizeof...(Args);
  }

  XLOIL_EXPORT FuncRegistrationMemo& createRegistrationMemo(const char* entryPoint_, size_t nArgs);

  template <class TFunc> inline FuncRegistrationMemo&
    registrationMemo(const char* name, TFunc func)
  {
    return createRegistrationMemo(name, getArgumentCount(func));
  }
 
  int
    findRegisteredFunc(const wchar_t* name, std::shared_ptr<FuncInfo>* info) noexcept;

  // TODO: registerAsync
  // TODO: registerMacro
}