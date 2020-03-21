#pragma once

#include "ExportMacro.h"
#include "StringUtils.h"

#ifdef _DEBUG
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_TRACE
#else
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_DEBUG
#endif 

#define SPDLOG_WCHAR_TO_UTF8_SUPPORT
#include <spdlog/spdlog.h> 
#include <string>


#define XLO_TRACE(...) SPDLOG_TRACE(__VA_ARGS__)
#define XLO_DEBUG(...) SPDLOG_DEBUG(__VA_ARGS__)
#define XLO_INFO(...) SPDLOG_INFO(__VA_ARGS__)
#define XLO_WARN(...) SPDLOG_WARN(__VA_ARGS__)
#define XLO_ERROR(...) SPDLOG_ERROR(__VA_ARGS__)

/// <summary>
/// Throws an xloil::Exception. Accepts python format strings like the logging functions
/// e.g. XLO_THROW("Bad: {0} {1}", errCode, e.what()) 
/// </summary>
#define XLO_THROW(...) do { throw xloil::Exception(__FILE__, __LINE__, __FUNCTION__, __VA_ARGS__); } while(false)
#define XLO_ASSERT(condition) assert((condition) && #condition)
//#define XLO_REQUIRES(condition) do { if (!(condition)) XLO_THROW("xloil requires: " #condition); } while(false)
//#define XLO_REQUIRES_MSG(condition, msg) do { if (!(condition)) XLO_THROW(msg); } while(false)


namespace xloil
{
  void initialiseLogger(const std::string& logLevel, const std::string* logFilePath);

  /// <summary>
  /// Wrapper around GetLastError and FormatMessage to write out any error condition
  /// set by Windows API functions.
  /// </summary>
  std::wstring writeWindowsError();

  /// <summary>
  /// Gets the logger registry for the core dll so plugins can output to the same
  /// log file
  /// </summary>
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry();

#pragma warning(disable: 4275) // Complaints about dll-interface. MS suggests disabling
  class Exception : public std::runtime_error
  {
  public:
    template<class... Args>
    Exception(const char* path, const int line, const char* func,
      std::string_view formatStr, Args &&... args)
      : Exception(path, line, func, formatMessage(formatStr, std::forward<Args>(args)...))
    {}
    template<class... Args>
    Exception(const char* path, const int line, const char* func,
      std::wstring_view formatStr, Args &&... args)
      : Exception(path, line, func, formatMessage(formatStr, std::forward<Args>(args)...))
    {}
    inline Exception(const char* path, const int line, const char* func, std::basic_string_view<wchar_t> msg)
      : Exception(path, line, func, utf16ToUtf8(msg.data()))
    {}

    XLOIL_EXPORT Exception(const char* path, const int line, const char* func, std::basic_string_view<char> msg);
    XLOIL_EXPORT virtual ~Exception() throw();

  private:
    int _line;
    std::string _file;
    std::string _function;

    template<class TChar, class... Args>
    std::basic_string<TChar> formatMessage(
      std::basic_string_view<TChar> formatStr,
      Args &&... args)
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