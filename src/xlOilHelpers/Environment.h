#pragma once

#include <string>

namespace xloil
{
  /// <summary>
  /// Helper function to capture C++ strings from Windows Api functions which have
  /// signatures like
  ///    int_charsWritten GetTheString(wchar* buffer, int bufferSize);
  /// </summary>
  namespace detail
  {
    template<class TChar, class F>
    auto captureStringBufferImpl(F bufWriter, size_t initialSize)
    {
      std::basic_string<TChar> s;
      s.resize(initialSize);
      size_t len;
      // We assume, hopefully correctly, that the bufWriter function on
      // failure returns either -1 or the required buffer length.
      while ((len = bufWriter(s.data(), s.length())) > s.length())
        s.resize(len == size_t(-1) ? s.size() * 2 : len);

      // However, some windows functions, e.g. ExpandEnvironmentStrings 
      // include the null-terminator in the returned buffer length whereas
      // other seemingly similar ones, e.g. GetEnvironmentVariable, do not.
      // Wonderful.
      s.resize(s.data()[len - 1] == '\0' ? len - 1 : len);
      return s;
    }
  }
  template<class F>
  auto captureStringBuffer(F bufWriter, size_t initialSize = 1024)
  {
    return detail::captureStringBufferImpl<char, F>(bufWriter, initialSize);
  }
  template<class F>
  auto captureWStringBuffer(F bufWriter, size_t initialSize = 1024)
  {
    return detail::captureStringBufferImpl<wchar_t, F>(bufWriter, initialSize);
  }

  /// <summary>
  /// Returns the value of specified environment variable
  /// or an empty string if it does not exist
  /// </summary>
  std::wstring getEnvVar(const wchar_t* name);
  std::string getEnvVar(const char * name);

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
    const wchar_t* hive,
    const wchar_t* location,
    std::wstring& result);

  bool getWindowsRegistryValue(
    const wchar_t* hive,
    const wchar_t* location,
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