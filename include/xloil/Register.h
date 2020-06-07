#pragma once
#include "ExportMacro.h"
#include <functional>
#include <vector>
#include <memory>
#include <list>

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
    /// Options which affect how the function is declared in Excel
    /// </summary>
    enum FuncOpts
    {
      /// <summary>
      /// The function is re-entrant and may be called ány of Excel's
      /// worker threads simultaneously
      /// </summary>
      THREAD_SAFE = 1 << 0,
      /// <summary>
      /// Gives the function special priviledges to read and write data 
      /// to the sheet
      /// </summary>
      MACRO_TYPE  = 1 << 1,
      /// <summary>
      /// Marks the function for recalculation on every calc cycle
      /// </summary>
      VOLATILE    = 1 << 2,
      /// <summary>
      /// Declares that the function is a command and has no return type
      /// </summary>
      COMMAND     = 1 << 3,
      /// <summary>
      /// Marks the function as asynchronous. Async functions do not return 
      /// directly but through a special handle which is passed as an argument
      /// </summary>
      ASYNC       = 1 << 4,
      /// <summary>
      /// Stops the function appearing in the function wizard or autocomplete
      /// </summary>
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

    /// <summary>
    /// Returns the number of function arguments 
    /// </summary>
    size_t numArgs() const { return args.size(); }
  };

  template<class TData> using RegisterCallbackT 
    = ExcelObj* (*)(TData* data, const ExcelObj**) noexcept;
  typedef RegisterCallbackT<void> RegisterCallback;

  template<class TData> using AsyncCallbackT 
    = void (*)(TData* data, const ExcelObj*, const ExcelObj**) noexcept;
  typedef AsyncCallbackT<void> AsyncCallback;

  using ExcelFuncObject = std::function<ExcelObj*(const FuncInfo& info, const ExcelObj**)>;
}