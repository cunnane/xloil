#pragma once
#include <xloil/ExportMacro.h>
#include <functional>
#include <vector>

namespace xloil { class ExcelObj; }

namespace xloil
{
  class ExcelObj;

  /// <summary>
  /// Holds the description of an Excel function argument
  /// </summary>
  struct FuncArg
  {
    enum ArgType
    {
      Obj         = 1 << 0,
      Range       = 1 << 1,
      Array       = 1 << 2,
      AsyncHandle = 1 << 3,
      ReturnVal   = 1 << 4,
      Optional    = 1 << 5 /// Just affects the auto generated help string
    };

    FuncArg(
      const wchar_t* argName = nullptr, 
      const wchar_t* argHelp = nullptr,
      const int argType = Obj)
      : name(argName ? argName : L"")
      , help(argHelp ? argHelp : L"")
      , type(argType)
    {}
    /// <summary>
    /// The name of the argument shown in the function wizard.
    /// </summary>
    std::wstring name;
    /// <summary>
    /// An optional help string for the argument displayed in the function wizard.
    /// </summary>
    std::wstring help;

    int type;

    bool operator==(const FuncArg& that) const
    {
      return name == that.name && help == that.help && type == that.type;
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
      /// Stops the function appearing in the function wizard or autocomplete
      /// </summary>
      HIDDEN      = 1 << 4
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
    unsigned options;

    /// <summary>
    /// Returns the number of function arguments 
    /// </summary>
    size_t numArgs() const { return args.size(); }
  };

  template<class TRet = ExcelObj*> using DynamicExcelFunc
    = std::function<TRet(const FuncInfo& info, const ExcelObj**)>;
}