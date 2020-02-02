#include "Log.h"
#include "Utils.h"
#include "WindowsSlim.h"
#include "Events.h"
#include "Interface.h"
#include "EntryPoint.h"
#include "spdlog/sinks/basic_file_sink.h"
#include <filesystem>
#include <algorithm>
namespace fs = std::filesystem;


namespace xloil
{
  namespace
  {
    std::string constructErrorString(const char* path, const int line, const char* func, const char* message)
    {
      std::string result(message);

      if (path)
      {
        const char* lastSlash = strrchr(path, '\\');
        if (lastSlash == 0)
          lastSlash = path;
        else
          ++lastSlash;
        const size_t pathLength = strlen(lastSlash);
        result += " (in ";
        std::transform(lastSlash, lastSlash + pathLength, std::back_inserter(result), tolower);
        result += ":" + std::to_string(line) + ")";
      }

      return result;
    }
  }

  XLOIL_EXPORT Exception::Exception(
    const char* path, const int line, const char* func, std::basic_string_view<char> msg)
    : runtime_error(constructErrorString(path, line, func, msg.data()).c_str())
  {}

  XLOIL_EXPORT Exception::~Exception() noexcept
  {}

  std::wstring writeWindowsError()
  {
    LPVOID lpMsgBuf;
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

    std::wstring result((wchar_t*)lpMsgBuf, size);
    LocalFree(lpMsgBuf);
    return result;
  }

  void initialiseLogger(const std::string& logLevel, const std::string* logFilePath)
  {
    auto logFile = logFilePath 
      ? *logFilePath 
      : wstring_to_utf8(fs::path(theXllPath()).replace_extension(".log").c_str());

    auto file_logger = spdlog::basic_logger_mt("basic_logger", logFile);

    // Flush on warnings or above
    file_logger->flush_on(spdlog::level::warn);
    spdlog::set_default_logger(file_logger);
    spdlog::set_level(spdlog::level::from_str(logLevel));

    // Flush log after each Excel calc cycle
    static auto handler = xloil::Event_CalcEnded() += [file_logger]() { file_logger->flush(); };
  }

  XLOIL_EXPORT spdlog::details::registry& loggerRegistry()
  {
    return spdlog::details::registry::instance();
  }
}