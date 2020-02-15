#pragma once
#include <string>
#include <codecvt>

#define XLO_STR(s) XLO_STR_IMPL(s)
#define XLO_STR_IMPL(s) #s

namespace xloil
{
  inline std::string wstring_to_utf8( const std::wstring_view& str)
  {
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    return converter.to_bytes(str.data());
  }

  inline std::wstring utf8_to_wstring(const std::string_view& str)
  {
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
    return converter.from_bytes(str.data());
  }
  template <class ...Args, template<class ...Args> class TContainer>
  inline auto utf8_to_wstring_v(const TContainer<Args...>& v)
  {
    return [v] {
      TContainer<std::wstring> tmp;
      tmp.reserve(v.size());
      std::transform(v.begin(), v.end(), back_inserter(tmp), utf8_to_wstring);
      return tmp;
    }();
  }
  template <class TInt> inline
  bool floatingToInt(double d, TInt& i)
  {
    double intpart;
    if (std::modf(d, &intpart) != 0.0)
      return false;

    // todo: ? std::numeric_limits<TInt>::
    if (!(intpart > INT_MIN && intpart < INT_MAX))
      return false;

    i = int(intpart);
    return true;
  }

  /// <summary>
  /// Helper function to capture C++ strings from Windows Api functions which have signatures like
  ///    int_charsWritten GetTheString(wchar* buffer, int bufferSize);
  /// </summary>
  template<class F>
  std::wstring captureWinApiString(F bufWriter, size_t initialSize = 1024)
  {
    std::wstring s;
    s.reserve(initialSize);
    size_t len;
    while ((len = bufWriter(s.data(), s.capacity())) > s.capacity())
      s.reserve(s.size() * 2);
    s._Eos(len);
    s.shrink_to_fit();
    return s;
  }

  /// <summary>
  /// Sets an environment variable, unsets when the object goes out of scope.
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
}