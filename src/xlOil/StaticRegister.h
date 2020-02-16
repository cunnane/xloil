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
    self& arg(const wchar_t* name, const wchar_t* help = nullptr)
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
    self& threadsafe()
    {
      _info->options |= FuncInfo::THREAD_SAFE;
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
}