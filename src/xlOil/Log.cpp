#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Events.h>
#include <xlOil/Interface.h>
#include <xloil/AppObjects.h>
#include <xloil/State.h>
#include <xlOil/Throw.h>
#include <xlOilHelpers/Exception.h>
#include "LogWindowSink.h"
#include <spdlog/sinks/basic_file_sink.h>
#include <spdlog/sinks/msvc_sink.h>
#include <spdlog/sinks/rotating_file_sink.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::wstring;
using std::string;
using std::make_shared;

namespace xloil
{
  namespace detail
  {
    void loggerInitialise(spdlog::level::level_enum level)
    {
      auto dbgWrite = make_shared<spdlog::sinks::msvc_sink_mt>();
      dbgWrite->set_level(level);

      auto logger = make_shared<spdlog::logger>("logger",
        spdlog::sinks_init_list{ dbgWrite });

      spdlog::initialize_logger(logger);

      // Flush on warnings or above
      logger->flush_on(spdlog::level::warn);
      spdlog::set_default_logger(logger);

      // Flush log after each Excel calc cycle
      static auto handler = Event::AfterCalculate() += [logger]() { logger->flush(); };
    }

    void loggerInitPopupWindow()
    {
      auto& state = App::internals();
      auto logWindow = makeLogWindowSink(
        (HWND)state.hWnd,
        (HINSTANCE)State::coreModuleHandle());

      auto logger = spdlog::default_logger();
      logger->sinks().push_back(logWindow);
    }

    void loggerAddFile(
      const wchar_t* logFilePath, const char* logLevel, 
      const size_t maxFileSizeKb, const size_t numFiles)
    {
      auto logger = spdlog::default_logger();
      auto fileWrite = make_shared<spdlog::sinks::rotating_file_sink_mt>(
        logFilePath, maxFileSizeKb * 1024, numFiles);
      fileWrite->set_level(spdlog::level::from_str(logLevel));
      logger->sinks().push_back(fileWrite);
      if (fileWrite->level() < logger->level())
        logger->set_level(fileWrite->level());
    }
  }

  XLOIL_EXPORT spdlog::details::registry& loggerRegistry()
  {
    return spdlog::details::registry::instance();
  }
}