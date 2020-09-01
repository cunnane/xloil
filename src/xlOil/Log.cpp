#include <xlOil/Log.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Events.h>
#include <xlOil/Interface.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/Throw.h>
#include <spdlog/sinks/basic_file_sink.h>
#include <spdlog/sinks/msvc_sink.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::wstring;
using std::string;
using std::make_shared;

namespace xloil
{
  XLOIL_EXPORT Exception::Exception(
    const char* path, 
    const int line, 
    const char* func, 
    const char* msg) noexcept
    : runtime_error(msg)
  {
    try
    {
      auto lastSlash = strrchr(path, '\\');
      XLO_INFO("{0} (in {2}:{3} during {1})",
        msg, func, lastSlash ? lastSlash + 1 : path, line);
    }
    catch (...)
    {}
  }

  std::wstring writeWindowsError()
  {
    wchar_t* lpMsgBuf = nullptr;
    auto dw = GetLastError();

    auto size = FormatMessage(
      FORMAT_MESSAGE_ALLOCATE_BUFFER |
      FORMAT_MESSAGE_FROM_SYSTEM |
      FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      dw,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
      (LPTSTR)&lpMsgBuf,
      0, NULL);

    auto msgBuf = std::shared_ptr<wchar_t>(lpMsgBuf, LocalFree);

    return wstring(msgBuf.get(), size);
  }

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

    void loggerAddFile(const wchar_t* logFilePath, const char* logLevel)
    {
      auto logger = spdlog::default_logger();
      auto fileWrite = make_shared<spdlog::sinks::basic_file_sink_mt>(
        utf16ToUtf8(logFilePath), false);
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