#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Events.h>
#include <xlOil/Throw.h>
#include <xlOil/State.h>
#include <xlOilHelpers/Exception.h>
#include "LogWindowSink.h"
#include <spdlog/sinks/basic_file_sink.h>
#include <spdlog/sinks/msvc_sink.h>
#include <spdlog/sinks/rotating_file_sink.h>

using std::wstring;
using std::string;
using std::make_shared;

namespace xloil
{
  std::shared_ptr<spdlog::logger> loggerInitialise(
    const char* debugStringLevel,
    bool makeDefault)
  {
    const auto debugWriterLevel = spdlog::level::from_str(debugStringLevel);

    auto logger = make_shared<spdlog::logger>("logger");

    if (debugWriterLevel != spdlog::level::off)
    {
      auto dbgWrite = make_shared<spdlog::sinks::msvc_sink_mt>();
      dbgWrite->set_level(debugWriterLevel);
      logger->sinks().push_back(dbgWrite);
    }

    spdlog::initialize_logger(logger);

    // Flush on warnings or above

    if (makeDefault)
      spdlog::set_default_logger(logger);

    return logger;
  }

  void loggerSetFlush(
    const std::shared_ptr<spdlog::logger>& logger,
    const char* flushLevel,
    bool flushAfterCalc)
  {
    const auto flushSpdLevel = spdlog::level::from_str(flushLevel);
    logger->flush_on(flushSpdLevel);

    if (flushAfterCalc)
      Event::AfterCalculate() += [logger]() { logger->flush(); };
  }

  void loggerAddPopupWindowSink(
    const std::shared_ptr<spdlog::logger>& logger)
  {
    auto& state = Environment::excelProcess();
    auto logWindow = makeLogWindowSink(
      (HWND)state.hWnd,
      (HINSTANCE)Environment::coreModuleHandle());

    logger->sinks().push_back(logWindow);
  }

  void loggerAddRotatingFileSink(
    const std::shared_ptr<spdlog::logger>& logger,
    const std::wstring_view& logFilePath, const char* logLevel,
    const size_t maxFileSizeKb, const size_t numFiles)
  {
    auto fileWrite = make_shared<spdlog::sinks::rotating_file_sink_mt>(
      wstring(logFilePath), maxFileSizeKb * 1024, numFiles);
    fileWrite->set_level(spdlog::level::from_str(logLevel));
    logger->sinks().push_back(fileWrite);
    if (fileWrite->level() < logger->level())
      logger->set_level(fileWrite->level());
  }
}