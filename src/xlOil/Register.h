#pragma once
#include "Options.h"
#include "ExportMacro.h"
#include <functional>
#include <vector>
#include <memory>
#include <list>



// Separate declaration needed to work around this quite serious MSVC compiler bug:
// https://stackoverflow.com/questions/45590594/generic-lambda-in-extern-c-function

/// Marks the start of an function regsistered in Excel
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

#define XLO_RETURN_ERROR(err) return ExcelObj::returnValue(err)

#define XLO_REGISTER_FUNC(func) extern auto _xlo_register_##func = xloil::registrationMemo(#func, func)


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
      , allowRange(false)
    {}
    /// <summary>
    /// The name of the argument shown in the function wizard.
    /// </summary>
    std::wstring name;
    /// <summary>
    /// An optional help string for the argument displayed in the function wizard.
    /// </summary>
    std::wstring help;

    /// <summary>
    /// If true, when the user specifies a cell reference in a function argument,
    /// Excel will pass a range, which xloil converts to an ExcelRange object.
    /// If false, cell references are converted to arrays or primitive types.
    /// </summary>
    bool allowRange;

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
      ASYNC       = 1 << 4,
      HIDDEN      = 1 << 5
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

  using ExcelFuncObject = std::function<ExcelObj*(const FuncInfo& info, const ExcelObj**)>;

  int
    findRegisteredFunc(const wchar_t* name, std::shared_ptr<FuncInfo>* info) noexcept;
}