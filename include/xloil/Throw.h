#include <spdlog/fmt/bundled/format.h>
#include <xlOilHelpers/StringUtils.h>
#include "ExportMacro.h"

#pragma once

/// <summary>
/// Throws an xloil::Exception. Accepts python format strings like the logging functions
/// e.g. XLO_THROW("Bad: {0} {1}", errCode, e.what()) 
/// </summary>
#define XLO_THROW(...) do { throw xloil::Exception(__FILE__, __LINE__, __FUNCTION__, __VA_ARGS__); } while(false)
#define XLO_ASSERT(condition) assert((condition) && #condition)

namespace xloil
{
  /// <summary>
  /// Wrapper around GetLastError and FormatMessage to write out any error condition
  /// set by Windows API functions.
  /// </summary>
  XLOIL_EXPORT std::wstring writeWindowsError();

#pragma warning(disable: 4275) // Complaints about dll-interface. MS suggests disabling
  class Exception : public std::runtime_error
  {
  public:
    template<class... Args>
    Exception(
      const char* path, const int line, const char* func,
      const std::string_view& formatStr, Args &&... args)
      : Exception(path, line, func, formatMessage(formatStr, std::forward<Args>(args)...))
    {}

    template<class... Args>
    Exception(
      const char* path, const int line, const char* func,
      const std::wstring_view& formatStr, Args &&... args)
      : Exception(path, line, func, formatMessage(formatStr, std::forward<Args>(args)...))
    {}

    inline Exception(
      const char* path, const int line, const char* func, 
      const std::wstring& msg)
      : Exception(path, line, func, utf16ToUtf8(msg.c_str()))
    {}

    inline Exception(
      const char* path, const int line, const char* func,
      const std::string& msg)
      : Exception(path, line, func, msg.c_str())
    {}

    XLOIL_EXPORT Exception(
      const char* path,
      const int line,
      const char* func,
      const char* message) noexcept;

  private:
    template<class TChar, class... Args>
    std::basic_string<TChar> formatMessage(
      std::basic_string_view<TChar> formatStr,
      Args &&... args) noexcept
    {
      try
      {
        // This 250 size is the same default as in spdlog, the buffer is actually dynamic
        fmt::basic_memory_buffer<TChar, 250> buf;
        fmt::format_to(buf, formatStr, std::forward<Args>(args)...);
        return std::basic_string<TChar>(buf.data(), buf.size());
      }
      catch (...)
      {
        return std::basic_string<TChar>(formatStr);
      }
    }
  };
}