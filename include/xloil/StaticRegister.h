#pragma once
#include "Register.h"
#include "ExcelObj.h"

namespace xloil { class FuncSpec; }

// In XLO_FUNC_START a separate declaration is needed to the function implementation
// to work around this quite serious MSVC compiler bug:
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
    self& optArg(const wchar_t* name, const wchar_t* help = nullptr, bool allowRange = false)
    {
      arg(name, help, allowRange);
      _info->args.back().optional = true;
      return *this;
    }
    self& arg(const wchar_t* name, const wchar_t* help = nullptr, bool allowRange = false)
    {
      _info->args.emplace_back(FuncArg(name, help));
      if (_allowRangeAll || allowRange)
        _info->args.back().allowRange = true;
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
    self& allowRange()
    {
      _allowRangeAll = true;
      return *this;
    }
    // TODO: public but not exported...can we hide this?
    std::shared_ptr<const FuncInfo> getInfo();

    std::string entryPoint;

  private:
    std::shared_ptr<FuncInfo> _info;
    bool _allowRangeAll;
    size_t _nArgs;
  };

  XLOIL_EXPORT FuncRegistrationMemo& 
    createRegistrationMemo(const char* entryPoint_, size_t nArgs);

  template <class R, class... Args> constexpr size_t
    countArguments(R(*)(Args...))
  {
    return sizeof...(Args);
  }

#ifndef _WIN64
  template <class R, class... Args> constexpr size_t
    countArguments(R(__stdcall *)(Args...))
  {
    return sizeof...(Args);
  }
#endif

  template <class TFunc> inline FuncRegistrationMemo&
    registrationMemo(const char* name, TFunc func)
  {
    return createRegistrationMemo(name, countArguments(func));
  }

  std::vector<std::shared_ptr<const FuncSpec>>
    processRegistryQueue(const wchar_t* moduleName);
}