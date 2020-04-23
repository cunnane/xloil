#include "StringUtils.h"
#include "WindowsSlim.h"


using std::wstring;

namespace xloil
{

  std::wstring getEnvVar(const wchar_t * name)
  {
    return captureStringBuffer(
      [name](auto* buf, auto len)
    {
      return GetEnvironmentVariable(name, buf, (DWORD)len);
    });
  }

  std::wstring expandEnvVars(const wchar_t* str)
  {
    return captureStringBuffer(
      [str](auto* buf, auto len)
    {
      return ExpandEnvironmentStrings(str, buf, (DWORD)len);
    }
    );
  }

  PushEnvVar::PushEnvVar(const wstring& name, const wstring& value)
    : PushEnvVar(name.c_str(), value.c_str())
  {}

  PushEnvVar::PushEnvVar(const wchar_t* name, const wchar_t* value)
    : _name(name)
    , _previous(getEnvVar(name))
  {
    auto s = expandEnvVars(value);
    SetEnvironmentVariable(name, s.c_str());
  }

  PushEnvVar::~PushEnvVar()
  {
    pop();
  }

  void PushEnvVar::pop()
  {
    if (_name.empty())
      return;

    SetEnvironmentVariable(_name.c_str(), _previous.c_str());
    _name.clear();
    _previous.clear();
  }
}