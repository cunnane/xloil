#pragma once

#include <xloil/ExportMacro.h>
#include <xloil/StringUtils.h>
#include <xloil/FmtStr.h>
#include <string>


#define XLO_LOGGER_CALL(level, ...) { xloil::Logger::instance().log(\
     xloil::Logger::Location{ __FILE__, __LINE__, __FUNCTION__ }, \
     level, XLO_FMT_((__VA_ARGS__))); }

#define XLO_TRACE(...) XLO_LOGGER_CALL(xloil::LogLevel::LOG_TRACE, __VA_ARGS__)
#define XLO_DEBUG(...) XLO_LOGGER_CALL(xloil::LogLevel::LOG_DEBUG, __VA_ARGS__)
#define XLO_INFO(...)  XLO_LOGGER_CALL(xloil::LogLevel::LOG_INFO,  __VA_ARGS__)
#define XLO_WARN(...)  XLO_LOGGER_CALL(xloil::LogLevel::LOG_WARN,  __VA_ARGS__)
#define XLO_ERROR(...) XLO_LOGGER_CALL(xloil::LogLevel::LOG_ERROR, __VA_ARGS__)

namespace xloil
{
  namespace detail
  {
    //void loggerInitialise(spdlog::level::level_enum level);
    void loggerInitPopupWindow();
    void loggerAddFile(
      const wchar_t* logFilePath, const char* logLevel, 
      size_t maxFileSizeKb, size_t numFiles = 1);
  }

  enum class LogLevel
  {
    LOG_TRACE = 0,
    LOG_DEBUG = 1,
    LOG_INFO = 2,
    LOG_WARN = 3,
    LOG_ERROR = 4,
    LOG_CRITICAL = 5,
    LOG_OFF = 6
  };

  class Logger
  {
  public:
    Logger();

    struct Location
    {
      const char* filename;
      int line;
      const char* funcname;
    };

    template<class Str>
    void log(Location location, LogLevel level, Str&& msg)
    {
      if (level < _minLevel)
        return;
      
      doLog(location, level, std::move(msg));
    }

    void setLevel(LogLevel level)
    {
      _minLevel = level;
    }

    XLOIL_EXPORT void doLog(
      const Location& location, 
      const LogLevel level, 
      std::wstring&& msg);

    XLOIL_EXPORT void doLog(
      const Location& location,
      const LogLevel level,
      std::string&& msg);

    XLOIL_EXPORT static Logger& instance();

  private:
    LogLevel _minLevel;
  };

}
