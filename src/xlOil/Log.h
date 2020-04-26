#pragma once

#include "ExportMacro.h"
#include <xlOilHelpers/StringUtils.h>

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
  void loggerInitialise(const char* logLevel);

  void loggerAddFile(const wchar_t* logFilePath);

  /// <summary>
  /// Gets the logger registry for the core dll so plugins can output to the same
  /// log file
  /// </summary>
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry();
}