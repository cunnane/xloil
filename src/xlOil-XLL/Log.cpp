#include <xloil/Log.h>
#define SPDLOG_WCHAR_TO_UTF8_SUPPORT
#include <xlOil/WindowsSlim.h>


#ifdef _DEBUG
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_TRACE
#else
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_DEBUG
#endif 


#include <spdlog/spdlog.h> 


namespace xloil
{
  Logger::Logger()
  {
    _minLevel = LogLevel::LOG_OFF;
  }
  
  void Logger::doLog(
    const Location& location,
    const LogLevel level,
    std::string&& msg)
  {
    spdlog::default_logger_raw()->log(
      spdlog::source_loc{ location.filename, location.line, location.funcname },
      spdlog::level::level_enum(level),
      msg);
  }

  void Logger::doLog(
    const Location& location,
    const LogLevel level,
    std::wstring&& msg)
  {
    spdlog::default_logger_raw()->log(
      spdlog::source_loc{ location.filename, location.line, location.funcname }, 
      spdlog::level::level_enum(level),
      msg);
  }

  Logger& Logger::instance()
  {
    static Logger theInstance; // TODO call logger initialise with level
    return theInstance;
  }
}