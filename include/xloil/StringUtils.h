#pragma once
#include <string>
#include <codecvt>
#include <algorithm>

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

    template <class TChar>
    struct char_traits : public std::char_traits<TChar>
    {};
    template <>
    struct char_traits<wchar_t> : public std::char_traits<wchar_t>
    {
      static wchar_t tolower(wchar_t c) { return ::towlower(c); }
    };
    template <>
    struct char_traits<char> : public std::char_traits<wchar_t>
    {
      static char tolower(char c) { return (char)::tolower(c); }
    };
  }
  /// <summary>
  /// Helper function to capture C++ strings from Windows Api functions which have
  /// signatures like
  ///    int_charsWritten GetTheString(wchar* buffer, int bufferSize);
  /// </summary>
  /// 
  template<class F>
  auto captureStringBuffer(F bufWriter, size_t initialSize = 1024)
  {
    return detail::captureStringBufferImpl<char, F>(bufWriter, initialSize);
  }

  /// <summary>
  /// Helper function to capture C++ strings from Windows Api functions which have
  /// signatures like
  ///    int_charsWritten GetTheString(wchar* buffer, int bufferSize);
  /// </summary>
  /// 
  template<class F>
  auto captureWStringBuffer(F bufWriter, size_t initialSize = 1024)
  {
    return detail::captureStringBufferImpl<wchar_t, F>(bufWriter, initialSize);
  }

  template <class Elem> inline
   std::basic_string<Elem>& toLower(std::basic_string<Elem>&& str)
  {
    std::transform(str.begin(), str.end(), str.begin(), detail::char_traits<Elem>::tolower);
    return str;
  }
  template <class Elem> inline
    void toLower(std::basic_string<Elem>& str)
  {
    std::transform(str.begin(), str.end(), str.begin(), detail::char_traits<Elem>::tolower);
  }


  /// <summary>
  /// Writes an unsigned int64 to a char buffer for a fixed radix. Does not
  /// write a null-terminator. Returns the length of the string written or
  /// zero if the buffer is insufficient.  Essentially `itoa` where the 
  /// compiler can optimise for the fixed radix.
  /// </summary>
  /// <typeparam name="TChar">char-type</typeparam>
  /// <typeparam name="TRadix">
  ///   Numbers above 32 will result in some pretty weird characters in the output
  /// </typeparam>
  /// <param name="value"></param>
  /// <param name="result">Pointer to the start of the result buffer</param>
  /// <param name="resultSize">Size of the result buffer in chars</param>
  /// <returns></returns>
  template<size_t TRadix, class TChar>
  inline uint8_t unsignedToString(size_t value, TChar* result, size_t resultSize)
  {
    // Surely larger than this is just silly?
    static_assert(TRadix <= 128);

    // 64 chars should hold a 64-bit int in radix = 2
    // TODO: could make this buffer smaller for bigger radix
    constexpr auto bufSize = 64;

    // The below division and remainder algorithm writes the string in reverse
    // so we write to a separate buffer
    TChar buf[bufSize];

    // Current write position in the buffer
    TChar* p = buf;
    do {
      // Integer divide value by radix 
      const uint8_t rem = value % TRadix;
      value = value / TRadix;

      // Remainder determines the offset from the '0' or 'a' chars
      if constexpr (TRadix <= 10)
      {
        *p = rem + '0';
      }
      else
      {
        *p = rem < 10
          ? rem + '0'
          : rem + 'a' - 10;
      }
      ++p;
    } while (value > 0);

    // Check how many chars we wrote
    const auto len = uint8_t(p - buf);

    // Check buffer size
    if (len > resultSize)
      return 0;

    // Reverse the string back into the result array
    std::reverse_copy(buf, p, result);
    return len;
  }

  template<size_t TRadix, class TChar, size_t TResultSize>
  inline size_t unsignedToString(size_t value, TChar(&result)[TResultSize])
  {
    return unsignedToString<TRadix, TChar>(value, result, TResultSize);
  }

  namespace detail
  {
    /// <summary>
    /// Parses a uint64 expressed as characters in a given alphabet.
    /// Stops parsing when `TAlphabet` assigns a value above `THigh`
    /// Moves the `begin` iterator to just past the last correctly
    /// parsed character.
    /// </summary>
    /// <typeparam name="TIter"></typeparam>
    /// <typeparam name="TAlphabet">
    ///   A trivially constructible class whose operator() returns 
    ///   a uint value for a given character 
    /// </typeparam>
    /// <typeparam name="TRadix">
    /// <typeparam name="THigh">
    ///   Maxmium allowed symbol value, usually TRadix - 1
    /// </typeparam>
    /// <param name="begin"></param>
    /// <param name="end"></param>
    /// <returns></returns>
    template<
      class TIter,
      class TAlphabet,
      size_t TRadix,
      size_t THigh = TRadix - 1>
    inline auto parseUnsigned(
      TIter& begin, 
      const TIter end)
    {
      size_t val = 0;
      do {
        auto c = *begin;
        auto v = TAlphabet()(c);
        if (v > THigh)
          break;
        val = val * TRadix + v;
      } while (++begin != end);
      return val;
    }

    struct StandardAlphabet
    {
      static constexpr int8_t _alphabet[] = {
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9,
        -1, -1, -1, -1, -1, -1, -1,
        10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
        -1, -1, -1, -1, -1, -1,
        10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
      };
      auto operator()(int8_t c) const 
      { 
        if (c < '0' || c > 'z')
          return (uint8_t)-1;
        return (uint8_t)_alphabet[c - '0'];
      }
    };

    struct DecimalAlphabet
    {
      auto operator()(int8_t c) const
      {
        if (c < '0' || c > '9')
          return (uint8_t)-1;
        return (uint8_t)(c - '0');
      }
    };
  }

  /// <summary>
  /// Parses a uint64 expressed as characters with a given radix.
  /// Essentially a templated version of `utoa`. Moves the `begin` 
  /// iterator to just past the last correctly parsed character.
  /// </summary>
  template<size_t TRadix, class TIter>
  inline auto parseUnsigned(TIter& begin, const TIter& end)
  {
    if constexpr (TRadix <= 10)
      return detail::parseUnsigned<TIter, detail::DecimalAlphabet, TRadix>(begin, end);
    else
      return detail::parseUnsigned<TIter, detail::StandardAlphabet, TRadix>(begin, end);
  }

  template<size_t TRadix, class TIter>
  inline auto parseUnsigned(const TIter& begin, const TIter& end)
  {
    auto i = begin;
    return parseUnsigned<TRadix, TIter>(i, end);
  }

  // Borrowed from Boost. Doesn't logically live in this header file, but 
  // lacks another home
  inline size_t boost_hash_combine(size_t seed) { return seed; }

  template <typename T, typename... Rest>
  inline size_t boost_hash_combine(size_t seed, const T& v, Rest... rest) {
    seed ^= std::hash<T>()(v) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
    return boost_hash_combine(seed, rest...);
  }

  template<class A, class B>
  struct pair_hash {
    size_t operator()(std::pair<A, B> p) const noexcept
    {
      return boost_hash_combine(377, p.first, p.second);;
    }
  };
}