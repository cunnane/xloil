#pragma once
#include "Register.h"

namespace xloil
{
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

  XLOIL_EXPORT FuncRegistrationMemo& createRegistrationMemo(const char* entryPoint_, size_t nArgs);

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
}