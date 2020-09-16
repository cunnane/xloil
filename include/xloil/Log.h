#pragma once

#include <xloil/ExportMacro.h>
#include <xlOil/StringUtils.h>
#include <xloil/Interface.h>

#ifdef _DEBUG
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_TRACE
#else
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_DEBUG
#endif 

#define SPDLOG_WCHAR_TO_UTF8_SUPPORT
#include <xlOil/WindowsSlim.h>
#include <spdlog/spdlog.h> 
#include <string>

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
    void loggerInitPopupWindow(const char* logLevel);
    void loggerAddFile(const wchar_t* logFilePath, const char* logLevel);
  }

  /// <summary>
  /// Gets the logger registry for the core dll so plugins can output to the same
  /// log file
  /// </summary>
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry();

  inline void linkLogger(AddinContext*, const PluginContext& plugin)
  {
    if (plugin.action == PluginContext::Load)
      spdlog::set_default_logger(loggerRegistry().default_logger());
  }
}