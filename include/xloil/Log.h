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
  /// <summary>
  /// Create an initialise a spdlog logger.
  /// <param name="debugLevel">
  ///   If set to a spdlog level other than "off", a sink which writes to `OutputDebugString`
  ///   is created with the given log level and added to the logger's sinks
  /// </param>
  /// <param name="makeDefault">
  ///   Make this logger the default logger, which can be accessed with `spdlog::default_logger()`
  /// </param>
  /// </summary>
  std::shared_ptr<spdlog::logger> 
    loggerInitialise(const std::string_view& debugLevel, bool makeDefault = true);

  /// <summary>
  /// Sets the log level at which the logger's sinks should be flushed (written to disk for
  /// file based sinks).
  /// </summary>
  void loggerSetFlush(
    const std::shared_ptr<spdlog::logger>& logger,
    const std::string_view& flushLevel);

  /// <summary>
  /// Adds a logger sink which pops up a log window when log messages exceed a certain
  /// threshold (which defaults to "warn") 
  /// </summary>
  void loggerAddPopupWindowSink(const std::shared_ptr<spdlog::logger>& logger);

  /// <summary>
  /// Add a rotating file sink to the logger, returning the name of the file
  /// which was created.  This may be different to the requested file if that
  /// file cannot be opened.
  /// </summary>
  std::wstring loggerAddRotatingFileSink(
    const std::shared_ptr<spdlog::logger>& logger,
    const std::wstring_view& logFilePath, const char* logLevel,
    size_t maxFileSizeKb, size_t numFiles = 1);

  /// <summary>
  /// Gets the logger registry for the core dll so plugins can output to the same
  /// log file
  /// </summary>
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry();
}

#endif