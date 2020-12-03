#pragma once
#include <string>
#include <codecvt>

namespace xloil
{
  /// <summary>
  /// Converts a UTF-16 wstring to a UTF-8 string
  /// </summary>
  inline std::string utf16ToUtf8(const std::wstring_view& str)
  {
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    return converter.to_bytes(str.data());
  }

  /// <summary>
  /// Converts a UTF-8 string to a UTF-16 wstring
  /// </summary>
  inline std::wstring utf8ToUtf16(const std::string_view& str)
  {
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
    return converter.from_bytes(str.data());
  }

  namespace detail
  {
    // http://unicode.org/faq/utf_bom.html
    constexpr char32_t LEAD_OFFSET = (char32_t)(0xD800 - (0x10000 >> 10));
    constexpr char32_t SURROGATE_OFFSET = (char32_t)(0x10000 - (0xD800 << 10) - 0xDC00);
    constexpr char32_t HI_SURROGATE_START = 0xD800;
  }

  /// <summary>
  /// Concerts a UTF-16 wchar_t string to a UTF-32 char32_t one.
  /// This string conversion appears to be missing from the standard codecvt
  /// library as of C++17.
  /// </summary>
  struct ConvertUTF16ToUTF32
  {
    using to_char = char32_t;
    using from_char = char16_t;

    size_t operator()(
      to_char* target, 
      const size_t targetSize,
      const from_char* begin, 
      const from_char* end) const noexcept
    {
      auto* p = target;
      auto* pEnd = target + targetSize;
      for (; begin < end; ++begin, ++p)
      {
        // If we are past the end of the buffer, carry on so we can
        // determine the required buffer length, but do not write
        // any characters
        if (p == pEnd)
        {
          if (*begin >= detail::HI_SURROGATE_START)
            ++begin;
        }
        else
        {
          if (*begin < detail::HI_SURROGATE_START)
            *p = *begin;
          else
          {
            auto lead = *begin++;
            *p = (lead << 10) + *begin + detail::SURROGATE_OFFSET;
          }
        }
      }
      return p - target;
    }
    size_t operator()(
      to_char* target, 
      const size_t size, 
      const wchar_t* begin, 
      const wchar_t* end) const noexcept
    {
      return (*this)(target, size, (const from_char*)begin, (const from_char*)end);
    }
  };

  struct ConvertUTF32ToUTF16
  {
    using from_char = char32_t;
    using to_char = char16_t;

    static void convertChar(char32_t codepoint, char16_t &h, char16_t &l) noexcept
    {
      if (codepoint < 0x10000)
      {
        h = (char16_t)codepoint;
        l = 0;
        return;
      }
      h = (char16_t)(detail::LEAD_OFFSET + (codepoint >> 10));
      l = (char16_t)(0xDC00 + (codepoint & 0x3FF));
    }

    /// <summary>
    /// Stops at null character
    /// </summary>
    /// <param name="target"></param>
    /// <param name="targetSize"></param>
    /// <param name="begin"></param>
    /// <param name="end"></param>
    /// <returns></returns>
    size_t operator()(
      to_char* target, 
      const size_t targetSize, 
      const from_char* begin, 
      const from_char* end) const noexcept
    {
      auto* p = target;
      auto* pEnd = target + targetSize;
      to_char lead, trail;
      for (; begin != end; ++begin, ++p)
      {
        convertChar(*begin, lead, trail);
        // If we are past the end of the buffer, carry on so we can
        // determine the required buffer length, but do not write
        // any characters
        if (p >= pEnd || (trail != 0 && p + 1 >= pEnd))
        {
          if (trail != 0) 
            ++p;
        }
        else
        {
          *p = lead;
          if (trail != 0)
            *(++p) = trail;
        }
      }
      return p - target;
    }
    size_t operator()(
      wchar_t* target, 
      const size_t size,
      const from_char* begin, 
      const from_char* end) const noexcept
    {
      return (*this)((to_char*)target, size, begin, end);
    }
  };

  /// <summary>
  /// strlen for char32 strings with a maximum length (in case the string
  /// is not null terminated). If a max is not required, use std::char_traits.
  /// </summary>
  /// <param name="str"></param>
  /// <param name="max"></param>
  /// <returns></returns>
  inline size_t strlen32(const char32_t* str, const size_t max)
  {
    size_t count = 0;
    while (*str != 0 && count < max)
    {
      ++count;
      ++str;
    }
    return count;
  }

  /// <summary>
  /// Tries to convert the provided floating point double to an integer.
  /// Returns false if the input has a fractional part or is too large for
  /// the given integer type.
  /// </summary>
  /// <param name="d"></param>
  /// <param name="i"></param>
  /// <returns></returns>
  template <class TInt> inline
  bool floatingToInt(double d, TInt& i) noexcept
  {
    double intpart;
    if (std::modf(d, &intpart) != 0.0)
      return false;

    if (!(intpart > (std::numeric_limits<TInt>::min)()
      && intpart < (std::numeric_limits<TInt>::max)()))
      return false;

    i = TInt(intpart);
    return true;
  }

  /// <summary>
  /// Wraps sprintf and returns a wstring
  /// </summary>
  template<class...Args>
  inline std::wstring
    formatStr(const wchar_t* fmt, Args&&...args)
  {
    const auto size = (size_t)_scwprintf(fmt, args...);
    std::wstring result(size + 1, 0);
    swprintf_s(&result[0], size + 1, fmt, args...);
    result.pop_back();
    return result;
  }

  /// <summary>
  /// Wraps sprintf and returns a string
  /// </summary>
  template<class...Args>
  inline std::string
    formatStr(const char* fmt, Args&&...args)
  {
    const auto size = (size_t)_scprintf(fmt, args...);
    std::string result(size + 1, 0);
    sprintf_s(&result[0], size + 1, fmt, args...);
    result.pop_back();
    return result;
  }

  template <class TChar>
  struct CaselessCompare
  {
    bool operator()(
      const std::basic_string<TChar> & lhs,
      const std::basic_string<TChar> & rhs) const
    {
      return (*this)(lhs.c_str(), rhs.c_str());
    }
    bool operator()(const TChar* lhs, const TChar* rhs) const
    {
      if constexpr (std::is_same<TChar, wchar_t>::value)
        return _wcsicmp(lhs, rhs) < 0;
      else
        return _stricmp(lhs, rhs) < 0;
    }
  };
}