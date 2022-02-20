#include <spdlog/fmt/bundled/format.h>
#include <xlOil/StringUtils.h>
#include "ExportMacro.h"

#pragma once

/// <summary>
/// Throws an xloil::Exception. Accepts python format strings like the logging functions
/// e.g. XLO_THROW("Bad: {0} {1}", errCode, e.what()) 
/// </summary>
#define XLO_THROW(...) do { throw xloil::Exception<>(__FILE__, __LINE__, __FUNCTION__, __VA_ARGS__); } while(false)
#define XLO_ASSERT(condition) assert((condition) && #condition)
#define XLO_THROW_TYPE(Type, ...) do { throw xloil::Exception<Type>(__FILE__, __LINE__, __FUNCTION__, __VA_ARGS__); } while(false)

namespace xloil
{
  /// <summary>
  /// Wrapper around GetLastError and FormatMessage to write out any error condition
  /// set by Windows API functions.
  /// </summary>
  XLOIL_EXPORT std::wstring writeWindowsError();

  XLOIL_EXPORT void logException(
    const char* path,
    const int line,
    const char* func,
    const char* msg) noexcept;

  namespace detail
  {
    inline std::string toUtf8(std::string&& str) { return std::move(str); }
    inline std::string toUtf8(std::wstring&& wstr) { return utf16ToUtf8(wstr); }

    template<class TChar, class... Args>
    inline std::string formatExceptionT(
      const std::basic_string_view<TChar>& formatStr,
      Args&&... args) noexcept
    {
      try
      {
        // This 250 size is the same default as in spdlog, the buffer is actually dynamic
        fmt::basic_memory_buffer<TChar, 250> buf;
        fmt::format_to(buf, formatStr, std::forward<Args>(args)...);
        return toUtf8(std::basic_string<TChar>(buf.data(), buf.size()));
      }
      catch (...)
      {
        return toUtf8(std::basic_string<TChar>(formatStr));
      }
    }
    template<class... Args>
    inline std::string formatException(
      const std::string_view& formatStr,
      Args&&... args) noexcept
    {
      return formatExceptionT<char>(formatStr, std::forward<Args>(args)...);
    }
    template<class... Args>
    inline std::string formatException(
      const std::wstring_view& formatStr,
      Args&&... args) noexcept
    {
      return formatExceptionT<wchar_t>(formatStr, std::forward<Args>(args)...);
    }

  } // detail

#pragma warning(disable: 4275) // Complaints about dll-interface. MS suggests disabling
  template<class TBase = std::runtime_error>
  class Exception : public TBase
  {
  public:
    template<class TMsg, class... Args>
    Exception(
      const char* path, const int line, const char* func,
      TMsg&& msg, Args&&... args)
      : TBase(detail::formatException(std::forward<TMsg>(msg), std::forward<Args>(args)...))
    {
      logException(path, line, func, what());
    }
  };
}