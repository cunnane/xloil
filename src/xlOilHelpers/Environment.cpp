#include "Environment.h"
#include <xloil/StringUtils.h>
#include <xloil/WindowsSlim.h>
#include "Exception.h"
#include <regex>
#include <cassert>

using std::wstring;
using std::wregex;
using std::wsmatch;
using std::wsregex_iterator;

namespace xloil
{
    std::wstring getEnvVar(const wchar_t * name)
    {
      return captureWStringBuffer(
        [name](auto* buf, auto len)
      {
        return GetEnvironmentVariableW(name, buf, (DWORD)len);
      });
    }

    std::string getEnvVar(const char * name)
    {
      return captureStringBuffer(
        [name](auto* buf, auto len)
      {
        return GetEnvironmentVariableA(name, buf, (DWORD)len);
      });
    }

    std::wstring expandEnvironmentStrings(const wchar_t* str)
    {
      return captureWStringBuffer(
        [str](auto* buf, auto len)
      {
        return ExpandEnvironmentStringsW(str, buf, (DWORD)len);
      });
    }

    PushEnvVar::PushEnvVar(const wstring& name, const wstring& value)
      : PushEnvVar(name.c_str(), value.c_str())
    {}

    PushEnvVar::PushEnvVar(const wchar_t* name, const wchar_t* value)
      : _name(name)
      , _previous(getEnvVar(name))
    {
      auto s = expandEnvironmentStrings(value);
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

    namespace
    {
      template<int RegType>
      bool getWindowsRegistryValue(
        const std::wstring& hive,
        const std::wstring& location,
        void* buffer,
        DWORD* bufSize)
      {
        HKEY root;
        if (hive == L"HKLM")
          root = HKEY_LOCAL_MACHINE;
        else if (hive == L"HKCU")
          root = HKEY_CURRENT_USER;
        else if (hive == L"HKCR")
          root = HKEY_CLASSES_ROOT;
        else
          return false;

        const auto lastSlash = location.rfind(L'\\');
        const auto subKey = location.substr(0, lastSlash);
        const auto value = lastSlash + 1 < location.size()
          ? location.substr(lastSlash + 1) : wstring();

        return ERROR_SUCCESS == RegGetValue(
          root,
          subKey.c_str(),
          value.c_str(),
          RegType,
          nullptr /*type not required*/,
          buffer,
          bufSize);
      }
    }

    bool getWindowsRegistryValue(
      const wchar_t* hive,
      const wchar_t* location,
      std::wstring& result)
    {
      assert(hive && location);
      wchar_t buffer[1024];
      DWORD bufSize = sizeof(buffer) / sizeof(wchar_t);
      if (getWindowsRegistryValue<RRF_RT_REG_SZ>(hive, location, buffer, &bufSize))
      {
        result = buffer;
        return true;
      }
      return false;
    }

    bool getWindowsRegistryValue(
      const wchar_t* hive,
      const wchar_t* location,
      unsigned long& result)
    {
      assert(hive && location);
      char buffer[sizeof(DWORD)];
      DWORD bufSize = sizeof(DWORD);
      if (getWindowsRegistryValue<RRF_RT_REG_DWORD>(hive, location, buffer, &bufSize))
      {
        result = *(DWORD*)buffer;
        return true;
      }
      return false;
    }

    static wregex registryExpander(L"<(HK[A-Z][A-Z])\\\\([^>]*)>",
      std::regex_constants::optimize | std::regex_constants::ECMAScript);

    wstring expandWindowsRegistryStrings(const std::wstring& str)
    {
      wstring result;

      wsregex_iterator next(str.begin(), str.end(), registryExpander);
      wsregex_iterator end;
      wsmatch match;
      wstring regValue;
      if (next == end)
        return str;
      while (next != end)
      {
        match = *next;
        assert(match.size() == 3);
        result += match.prefix().str();
        if (getWindowsRegistryValue(match[1].str().c_str(), match[2].str().c_str(), regValue))
          result += regValue;
        next++;
      }
      result += match.suffix().str();

      return result;
    }

    namespace Helpers
    {
      std::wstring writeWindowsError()
      {
        wchar_t* lpMsgBuf = nullptr;
        const auto dw = GetLastError();
        const auto size = FormatMessage(
          FORMAT_MESSAGE_ALLOCATE_BUFFER |
          FORMAT_MESSAGE_FROM_SYSTEM |
          FORMAT_MESSAGE_IGNORE_INSERTS,
          NULL,
          dw,
          MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
          (LPTSTR)&lpMsgBuf,
          0, NULL);

        // Frees the buffer on exit
        const auto msgBuf = std::unique_ptr<wchar_t, decltype(&LocalFree)>(
          lpMsgBuf, LocalFree);

        return wstring(lpMsgBuf, size);
      }

      std::pair<HANDLE, std::wstring> makeTempFile()
      {
        wchar_t szTempFileName[MAX_PATH];
        wchar_t lpTempPathBuffer[MAX_PATH];

        //  Gets the temp path env string (no guarantee it's a valid path).
        auto retVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
        if (retVal > MAX_PATH || retVal == 0)
          throw Exception("GetTempPath failed");

        retVal = GetTempFileName(lpTempPathBuffer,
          L"XLO", // prefix 
          0,      // use system time for uniquifier
          szTempFileName);
        if (retVal == 0)
          throw Exception("GetTempFileName failed");

        //  Creates the new file to write to for the upper-case version.
        auto handle = CreateFile((LPTSTR)szTempFileName, // file name 
          GENERIC_READ,        // open for write 
          FILE_SHARE_READ | FILE_SHARE_WRITE,     // share everything
          NULL,                 // default security 
          CREATE_ALWAYS,        // overwrite existing
          FILE_ATTRIBUTE_TEMPORARY, //| FILE_FLAG_DELETE_ON_CLOSE,
          NULL);                // no template 

        if (handle == INVALID_HANDLE_VALUE)
          throw Exception("CreateFile failed");

        return std::make_pair(handle, wstring(szTempFileName));
      }
    }
  }