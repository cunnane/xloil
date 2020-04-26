#include "Log.h"
#include <xlOilHelpers/StringUtils.h>
#include <xlOilHelpers/WindowsSlim.h>
#include "Events.h"
#include "Interface.h"
#include "EntryPoint.h"
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
    const char* path, const int line, const char* func, std::basic_string_view<char> msg)
    : runtime_error(msg.data())
    , _line(line)
    , _file(path)
    , _function(func)
  {
    XLO_ERROR("{0} (in {2}:{3} during {1})", msg.data(), func, fs::path(path).filename().string(), line);
  }

  XLOIL_EXPORT Exception::~Exception() noexcept
  {}

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

  void loggerInitialise(const char* logLevel)
  {
    auto dbgWrite = make_shared<spdlog::sinks::msvc_sink_mt>();
    
    auto logger = make_shared<spdlog::logger>("logger", 
      spdlog::sinks_init_list{ dbgWrite });

    spdlog::initialize_logger(logger);
    // Flush on warnings or above
    logger->flush_on(spdlog::level::warn);
    spdlog::set_default_logger(logger);
    spdlog::set_level(spdlog::level::from_str(logLevel));
    
    // Flush log after each Excel calc cycle
    static auto handler = xloil::Event_CalcEnded() += [logger]() { logger->flush(); };
  }

  void loggerAddFile(const wchar_t* logFilePath)
  {
    auto fileWrite = make_shared<spdlog::sinks::basic_file_sink_mt>(
      utf16ToUtf8(logFilePath), false);
    spdlog::default_logger()->sinks().push_back(fileWrite);
  }

  XLOIL_EXPORT spdlog::details::registry& loggerRegistry()
  {
    return spdlog::details::registry::instance();
  }
}