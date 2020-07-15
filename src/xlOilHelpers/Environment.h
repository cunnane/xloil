#pragma once

#include <string>

namespace xloil
{
  /// <summary>
  /// Helper function to capture C++ strings from Windows Api functions which have signatures like
  ///    int_charsWritten GetTheString(wchar* buffer, int bufferSize);
  /// </summary>
  template<class F>
  std::wstring captureStringBuffer(F bufWriter, size_t initialSize = 1024)
  {
    std::wstring s;
    s.reserve(initialSize);
    size_t len;
    // We assume, hopefully correctly, that the bufWriter function on
    // failure returns either -1 or the required buffer length.
    while ((len = bufWriter(s.data(), s.capacity())) > s.capacity())
      s.reserve(len == size_t(-1) ? s.size() * 2 : len);

    // However, some windows functions, e.g. ExpandEnvironmentStrings 
    // include the null-terminator in the returned buffer length whereas
    // other seemingly similar ones, e.g. GetEnvironmentVariable, do not.
    // Wonderful.
    s._Eos(s.data()[len - 1] == '\0' ? len - 1 : len);
    s.shrink_to_fit();
    return s;
  }

  /// <summary>
  /// Returns the value of specified environment variable
  /// or an empty string if it does not exist
  /// </summary>
  std::wstring getEnvVar(const wchar_t* name);

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
  /// Gets a string value from the registry give a hive and key.
  /// Hive can be HKCU, HKCR, HKLM. A trailing slash can be used 
  /// to fetch the default value for that key.
  /// 
  /// Returns false if the key was not matched.
  /// </summary>
  bool getWindowsRegistryValue(
    const std::wstring& hive,
    const std::wstring& location,
    std::wstring& result);

  bool getWindowsRegistryValue(
    const std::wstring& hive,
    const std::wstring& location,
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
}