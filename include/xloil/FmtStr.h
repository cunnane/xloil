#pragma once
#include <string>
#include <array>

/// <summary>
/// Takes a format *string literal* of char or wchar_t and arguments to insert, 
/// similar to printf.
/// 
/// The format string must have the form:
/// 
///   "Deduced type Arg1: {}, Specifed type Arg2: {<printf specifed>}"
/// 
/// **At compile time** type deduction takes place and specified format 
/// strings are validated against the argument type. A printf-style format
/// string is generated which is passed to printf at runtime.
/// 
/// Supported argument types are:
/// 
///   * numeric types (incl bool)
///   * void*
///   * null-terminated c-strings: char*, wchar_t*
///   * STL strings: std::string, std::wstring
/// 
/// </summary>
#define XLO_FMT_1(format, ...) [&] () { \
  struct Str { static constexpr auto str() { return format; } }; \
  return xloil::detail::fmtstr(Str{}, __VA_ARGS__); \
}()
#define XLO_FMT_0(format) xloil::detail::fmtstr(format)

#define XLO_FMT_SELECT(_1,_2,_3,_4,_5,_6,_7,_8,_9,_10,_11,_12,_13,_14,_15, N,...) N
#define XLO_FMT_SELECT_(args) XLO_FMT_SELECT args
#define __XLO_CONCAT(a, b) a##b
#define _XLO_CONCAT(a, b) __XLO_CONCAT(a, b)
#define XLO_CONCAT(a, b) _XLO_CONCAT(a, b)
#define XLO_FMT(...) XLO_CONCAT(XLO_FMT_, XLO_FMT_SELECT_((__VA_ARGS__, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0))(__VA_ARGS__))
#define XLO_FMT_(args) XLO_FMT args

namespace xloil
{
  namespace detail
  {
    /// Holder for a C-style array which keeps track of the size. We do not
    /// use std::array as we want the type to be homogenous regardless of size
    /// </summary>
    template<class T>
    struct CArray
    {
      const T* data;
      size_t size;

      constexpr CArray() : data(nullptr), size(0) {}
      template<size_t N> constexpr CArray(const T(&arr)[N]) : data(arr), size(N) {}

      constexpr const T* begin() const { return data; }
      constexpr const T* end() const { return data + size; }
    };


    /// <summary>
    /// Class whose `value` is an array of valid printf format strings for the 
    /// give type 'T'. For numeric types, more than one format string can be
    /// valid.
    /// </summary>
    template <typename T> struct FormatSpec;

#define XLO_FMT_SPECIFIER(Type, ...) template<> struct FormatSpec<Type> \
    { static constexpr const char* value[] = { __VA_ARGS__}; }

    XLO_FMT_SPECIFIER(int, "d");
    XLO_FMT_SPECIFIER(long int, "ld");
    XLO_FMT_SPECIFIER(long long int, "lld");
    XLO_FMT_SPECIFIER(unsigned, "u", "o", "x");
    XLO_FMT_SPECIFIER(unsigned long int, "lu", "lo", "lx");
    XLO_FMT_SPECIFIER(unsigned long long int, "llu", "llo", "llx");
    XLO_FMT_SPECIFIER(char, "c", "d");
    XLO_FMT_SPECIFIER(signed char, "hhd");
    XLO_FMT_SPECIFIER(wchar_t, "lc");
    XLO_FMT_SPECIFIER(short, "d");
    XLO_FMT_SPECIFIER(unsigned short, "hu");
    XLO_FMT_SPECIFIER(unsigned char, "hhu");
    XLO_FMT_SPECIFIER(bool, "d");
    XLO_FMT_SPECIFIER(float, "f", "g", "e");
    XLO_FMT_SPECIFIER(double, "f", "g", "e");
    XLO_FMT_SPECIFIER(long double, "Lf");

    XLO_FMT_SPECIFIER(void*, "p");
    XLO_FMT_SPECIFIER(const void*, "p");
    XLO_FMT_SPECIFIER(char*, "s");
    XLO_FMT_SPECIFIER(const char*, "s");
    XLO_FMT_SPECIFIER(std::string, "s");
    XLO_FMT_SPECIFIER(wchar_t*, "ls");
    XLO_FMT_SPECIFIER(const wchar_t*, "ls");
    XLO_FMT_SPECIFIER(std::wstring, "ls");
#undef XLO_FMT_SPECIFIER


    /// <summary>
    /// Looks for character 'what' in string str, returning ptr to
    /// it if found else null.
    /// </summary>
    template<class TCharL, class TCharR>
    constexpr auto find_next_chr(const TCharL* str, TCharR what)
    {
      for (; *str != 0; ++str)
        if (*str == what) return str;
      
      return (const TCharL*)0;
    }

    /// <summary>
    /// std::copy with constexpr (not required in c++20)
    /// </summary>
    template <class TInIt, class TOutIt>
    constexpr auto iter_copy(TInIt first, TInIt last, TOutIt dest)
    {
      for (; first != last; ++dest, ++first)
        *dest = *first;
      
      return dest;
    }

    // TODO: replace with copy_while
    template<class TChar, class TIter>
    constexpr auto str_copy(const TChar* source, TIter dest)
    {
      auto len = std::char_traits<TChar>::length(source);
      return iter_copy(source, source + len, dest);
    }

    /// <summary>
    /// Copies source iterator to destination iterator whilst predicate(*source) 
    /// is true.  It would seem more STL-like to return a pair of updated iterators 
    /// (src, dest), but tuple unpacking isn't constexpr in C++17, so we rather
    /// mutate the iterator arguments.
    /// </summary>
    template <class TInIt, class TOutIt, class TPred>
    constexpr auto copy_while(TInIt& source, TOutIt& dest, TPred predicate)
    {
      for (; predicate(*source); ++dest, ++source)
        *dest = *source;
    }

    /// <summary>
    /// Essentially constexpr version of `strncmp(left, right, nChars) != 0`
    /// </summary>
    template<class TCharL, class TCharR>
    constexpr bool str_match_n(const TCharL* left, const TCharR* right, size_t n_chars)
    {
      for (size_t i = 0; i < n_chars; ++i)
        if (*left++ != *right++)
          return false;
      return true;
    }

    /// <summary>
    /// Creates an array containing the valid format strings for each arg
    /// </summary>
    template<class... Args>
    struct ArgFormats
    {
      static constexpr auto N = sizeof...(Args);
      std::array<CArray<const char*>, N> formats;

      constexpr ArgFormats()
      {
        // Is it possible to initialise directly with initialiser list?
        auto i = 0;
        (void(formats[i++] = CArray<const char*>(FormatSpec<std::decay_t<Args>>::value)), ...);
      }

      constexpr auto get(size_t iArg, size_t i = 0) const
      {
        // use first argfmt by default
        return formats[iArg].data[i];
      }

      /// <summary>
      /// Returns true if the end chars of the provided string matches
      /// one of the allowed formats for the given arg index.
      /// Expects `str` to point to the *end* of the format string
      /// </summary>
      template<class TChar>
      constexpr auto suffix_match(size_t iArg, const TChar* str) const
      {
        for (auto i = 0; i < formats[iArg].size; ++i)
        {
          auto fmt = formats[iArg].data[i];
          auto len = std::char_traits<char>::length(fmt);
          if (str[-1] == fmt[len - 1])
            return i;
        }
        return -1;
      }
    };

    /// <summary>
    /// Iterator which keeps track of position but ignores assgnment. Used to 
    /// determine required string length. May exist in STL somwhere... 
    /// </summary>
    struct CountingIterator
    {
      struct value_type {
        template<class T> constexpr auto operator=(T) noexcept { return *this; }
      };
      constexpr auto operator++()    noexcept { ++count; return *this; }
      constexpr auto operator++(int) noexcept { ++count; return *this; }
      constexpr auto operator*()     noexcept { return value_type(); }
      constexpr auto operator=(CountingIterator that) noexcept { count = that.count; return *this; }
      constexpr auto operator-(CountingIterator that) const noexcept { return count - that.count; }
      size_t count = 0;
    };

    template<class TChar, class TIter, class TFmts>
    constexpr auto write_fmtstr(const TChar* str, TIter destination, TFmts arg_fmts)
    {
      auto dest = destination;
      auto source = str;

      // Copy up to the first '{' or the end of str
      copy_while(source, dest, [](auto c) { return c != '{' && c != 0; });

      // For each arg:
      //   * Source iterator should point at a '{'
      //   * Find next '}'
      //   * If a format string is found between the braces, check it matches
      //     the type of the corresponding argument
      //   * Otherwise insert the first appropriate format string
      //   * Copy string to next '{'
      // 
      // The 'throw' calls below don't actually produce a nice error message but 
      // rather cause constexpr passing to fail on that line.
      // 
      // TODO: support *, .* (variable precision, need to eat the next int argument)
      // TODO: support precision specifier with no arg type, e.g. {.6}
      // TODO: escape {

      auto i_arg = 0;
      while (*source != 0)
      {
        if (*source != '{') throw "Not enough braces";
        ++source; // eat the brace

        auto close_brace = find_next_chr(source, '}');
        if (!close_brace) throw "No close brace";

        (*dest++) = (TChar)'%';

        // Has format specifier been provided?
        if (close_brace - source > 0)
        {
          auto prev_src = source;
          auto arg_num = parseUnsigned<10>(source, close_brace);
          if (source == prev_src)
            arg_num = i_arg++;
          else if (i_arg > 0)
            throw "Mismatched arg nums";
          else
            i_arg = -666;

          if (arg_num >= arg_fmts.N)
            throw "Arg num out of range";

          if (*source != ':')
            dest = str_copy(arg_fmts.get(arg_num), dest);
          else
          {
            ++source; // eat colon
            auto i_fmt = arg_fmts.suffix_match(arg_num, close_brace);
            if (i_fmt < 0)
              throw "Format specifier doesn't match type";
            dest = iter_copy(source, close_brace - 1, dest);
            dest = str_copy(arg_fmts.get(arg_num, i_fmt), dest);
          }
        }
        else
        {
          dest = str_copy(arg_fmts.get(i_arg), dest);
          ++i_arg;
        }

        source = close_brace + 1;
        copy_while(source, dest, [](auto c) { return c != '{' && c != 0; });
      }

      // Add terminator
      (*dest++) = (TChar)0;

      return dest - destination; // std::distance(destination, dest);
    }

    /// <summary>
    /// Creates (and returns) a std::array<char_type> of the appropriate
    /// size, populated with the transformed format string.
    /// </summary>
    template<class TStr, typename... Args>
    constexpr auto build_fmtstr()
    {
      using char_type = typename std::decay_t<decltype(*TStr::str())>;
      constexpr ArgFormats<Args...> arg_fmts{};
      constexpr auto length = write_fmtstr(TStr::str(), CountingIterator(), arg_fmts);
      std::array<char_type, length> result{ 0 };
      write_fmtstr(TStr::str(), result.begin(), arg_fmts);
      return result;
    }

    template<typename...Args>
    inline auto call_sprintf(char* buf, Args... args)
    {
      return snprintf(buf, args...);
    }

    template<typename...Args>
    inline auto call_sprintf(wchar_t* buf, Args... args)
    {
      return swprintf(buf, args...);
    }

    template<class T>
    struct ReplaceString
    {
      auto operator()(T x) const { return x; }
    };

    template<class TChar>
    struct ReplaceString<std::basic_string<TChar>>
    {
      auto operator()(const std::basic_string<TChar>& x) const { return x.c_str(); }
    };

    template <class T> inline auto guessSize(T) { return 10; }
    //inline auto guessSize(unsigned long long) { return 10; }
    //inline auto guessSize(long long) { return 10; }
    //inline auto guessSize(void*) { return 10; }
    //inline auto guessSize(double) { return 10; }
    template<class TChar> inline auto guessSize(const std::basic_string<TChar>& s) { return s.length(); }
    template<class TChar> inline auto guessSize(const TChar* s) { return std::char_traits<TChar>::length(s); }


    template<class TChar, class... Args>
    inline auto stringprintf(const TChar* fmt, const size_t len_fmt, Args... args)
    {
      std::basic_string<TChar> result;

      // We can use _scprintf to get the correct number of chars but it's probably
      // more performant to just take a guess
      const auto size = len_fmt + (guessSize(args) + ...);
      result.resize(size);

      int charsWritten;
      while (true)
      {
        charsWritten = call_sprintf(
          &result[0],
          result.size(),
          fmt,
          ReplaceString<Args>()(args)...);
        if (charsWritten >= 0)
          break;

        if (errno == EOVERFLOW)
          result.resize(result.size() * 2);
        else
          throw std::runtime_error("sprintf badness");
      }
      result.resize(charsWritten);

      return std::move(result);
    }

    /// <summary>
    /// Takes a literal format string, provided by the constexpr return value of 
    /// TStr::str(), transforms it to printf style and invokes sprintf. See `XLO_FMT`. 
    /// </summary>
    template<class TStr, typename...Args>
    inline auto fmtstr(TStr, Args... args)
    {
      static constexpr auto printf_fmt = build_fmtstr<TStr, Args...>();
      return stringprintf(printf_fmt.data(), printf_fmt.size(), std::forward<Args>(args)...);
    }

    template<class TChar>
    inline auto fmtstr(const TChar* str)
    {
      return std::basic_string<TChar>(str);
    }
    template<class TChar>
    inline auto fmtstr(std::basic_string<TChar> str)
    {
      return str;
    }
  }
}