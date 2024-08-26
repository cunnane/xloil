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
#include <filesystem>

using std::wstring;
using std::string;
using std::make_shared;
namespace fs = std::filesystem;

namespace xloil
{
  std::shared_ptr<spdlog::logger> loggerInitialise(
    const std::string_view& debugLevel,
    bool makeDefault)
  {
    const auto debugWriterLevel = spdlog::level::from_str(string(debugLevel));

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
    const std::string_view& flushLevel)
  {
    const auto flushSpdLevel = spdlog::level::from_str(string(flushLevel));
    logger->flush_on(flushSpdLevel);
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

  std::wstring loggerAddRotatingFileSink(
    const std::shared_ptr<spdlog::logger>& logger,
    const std::wstring_view& logFilePath, const char* logLevel,
    const size_t maxFileSizeKb, const size_t numFiles)
  {
    auto filename = wstring(logFilePath);

    // Open for exclusive acces to check if another Excel instance is using the log file
    auto handle = CreateFile(
      filename.c_str(), 
      GENERIC_READ,
      0, // no sharing, exclusive
      NULL, OPEN_ALWAYS, 0, NULL);

    if ((handle != NULL) && (handle != INVALID_HANDLE_VALUE))
      CloseHandle(handle);
    else
      filename = fs::path(filename).replace_extension(std::to_wstring(GetCurrentThreadId()) + L".log");

    auto fileWrite = make_shared<spdlog::sinks::rotating_file_sink_mt>(
      filename, maxFileSizeKb * 1024, numFiles);
    fileWrite->set_level(spdlog::level::from_str(logLevel));
    logger->sinks().push_back(fileWrite);

    if (fileWrite->level() < logger->level())
      logger->set_level(fileWrite->level());

    return fileWrite->filename();
  }
}