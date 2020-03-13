#include "Utils.h"
#include "WindowsSlim.h"


using std::wstring;

namespace xloil
{
  PushEnvVar::PushEnvVar(const wstring& name, const wstring& value)
    : PushEnvVar(name.c_str(), value.c_str())
  {}

  PushEnvVar::PushEnvVar(const wchar_t* name, const wchar_t* value)
    : _name(name)
    , _previous(
        captureStringBuffer(
          [name](auto* buf, auto len) 
          { 
            return GetEnvironmentVariable(name, buf, (DWORD)len); 
          }
        )
      )
  {
    auto s = captureStringBuffer(
      [value](auto* buf, auto len) 
      { 
        return ExpandEnvironmentStrings(value, buf, (DWORD)len);
      }
    );
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