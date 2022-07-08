#pragma once

// TODO: slim down this header!

#ifdef XLOIL_NO_SPDLOG

#define XLO_TRACE(...) 
#define XLO_DEBUG(...) 
#define XLO_INFO(...) 
#define XLO_WARN(...) 
#define XLO_ERROR(...) 

#else

#include <xloil/ExportMacro.h>
#include <string>

#ifdef _DEBUG
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_TRACE
#else
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_DEBUG
#endif 

#define SPDLOG_WCHAR_FILENAMES
#define SPDLOG_WCHAR_TO_UTF8_SUPPORT
#include <xlOil/WindowsSlim.h>
#include <spdlog/spdlog.h> 


#define XLO_TRACE(...) SPDLOG_TRACE(__VA_ARGS__)
#define XLO_DEBUG(...) SPDLOG_DEBUG(__VA_ARGS__)
#define XLO_INFO(...) SPDLOG_INFO(__VA_ARGS__)
#define XLO_WARN(...) SPDLOG_WARN(__VA_ARGS__)
#define XLO_ERROR(...) SPDLOG_ERROR(__VA_ARGS__)

namespace xloil
{
  namespace detail
  {
    void loggerInitialise(spdlog::level::level_enum level);
    void loggerInitPopupWindow();
    void loggerAddFile(
      const wchar_t* logFilePath, const char* logLevel, 
      size_t maxFileSizeKb, size_t numFiles = 1);
  }

  /// <summary>
  /// Gets the logger registry for the core dll so plugins can output to the same
  /// log file
  /// </summary>
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry();
}

#endif