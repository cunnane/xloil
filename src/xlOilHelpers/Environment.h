#pragma once
#include <string>

typedef void* HANDLE;

namespace xloil
{
  /// <summary>
  /// Returns the value of specified environment variable
  /// or an empty string if it does not exist
  /// </summary>
  std::wstring getEnvironmentVar(const wchar_t* name);
  std::string getEnvironmentVar(const char * name);


  /// <summary>
  /// Sets the enviroment variable to a given value, returning 
  /// false if the action fails
  /// </summary>
  bool setEnvironmentVar(const wchar_t* name, const wchar_t* value);
  bool setEnvironmentVar(const char* name, const char* value);

  /// <summary>
  /// Expands environment variables in the specified string.
  /// Equivalent to the Windows API ExpandEnvironmentStrings
  /// </summary>
  std::wstring expandEnvironmentStrings(const wchar_t* str);
  inline std::wstring expandEnvironmentStrings(const std::wstring& str)
  {
    return expandEnvironmentStrings(str.c_str());
  }

  /// <summary>
  /// Sets an environment variable and unsets when the object goes out of scope.
  /// </summary>
  class PushEnvVar
  {
  private:
    std::wstring _previous;
    std::wstring _name;

  public:
    PushEnvVar(const std::wstring& name, const std::wstring& value);
    PushEnvVar(const wchar_t* name, const wchar_t* value);
    ~PushEnvVar();
    void pop();
  };

  /// <summary>
  /// Calls SetDllDirectory on the provided path and restores the previous
  /// value in it's dtor.  Safer than calling SetDllDirectory(NULL) which
  /// resets the search path.
  /// </summary>
  class PushDllDirectory
  {
  private:
    wchar_t _previous[260]; // MAX_PATH = 260
  public:
    PushDllDirectory(const wchar_t* path);
    PushDllDirectory(const char* path);
    ~PushDllDirectory();
  };

  /// <summary>
  /// Gets a string value from the registry give a hive and key.
  /// Hive can be HKCU, HKCR, HKLM. A trailing slash can be used 
  /// to fetch the default value for that key.
  /// 
  /// Returns false if the key was not matched.
  /// </summary>
  bool getWindowsRegistryValue(
    const std::wstring_view& hive,
    const std::wstring_view& location,
    std::wstring& result);

  bool getWindowsRegistryValue(
    const std::wstring_view& hive,
    const std::wstring_view& location,
    unsigned long& result);

  /// <summary>
  /// Matches and expands registry keys in the given string. Registry
  /// keys should be specified in the form "<(HKXX)\(Reg\Key\Value)>" 
  /// where HKXX can be HKCU, HKCR, HKLM. Only string values can be
  /// matched.
  /// 
  /// A trailing slash can be used to fetch the default value for that
  /// key.
  /// 
  /// Behaves similarly to expandEnvironmentStrings in that unmatched
  /// keys will be replaced with an empty string
  /// </summary>
  std::wstring expandWindowsRegistryStrings(const std::wstring& str);

  namespace Helpers
  {
    std::pair<HANDLE, std::wstring> makeTempFile();
  }
}