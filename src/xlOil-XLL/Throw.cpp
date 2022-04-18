#include <xloil/Throw.h>
#include <xloil/Log.h>
#include <xlOilHelpers/Exception.h>

namespace xloil
{
  XLOIL_EXPORT void logException(
    const char* path,
    const int line,
    const char* func,
    const char* msg) noexcept
  {
    try
    {
      auto lastSlash = strrchr(path, '\\');
      auto filename = lastSlash ? lastSlash + 1 : path;
      XLO_FMT("{0} (in {2}:{3} during {1})",
        msg, func, filename, line);

      {
        auto xlo_msg = XLO_FMT("{0} (in {2}:{3} during {1})",
          msg, func, filename, line);
        xloil::Logger::instance().log(
          xloil::Logger::Location{ __FILE__, __LINE__, __FUNCTION__ }, 
          xloil::LogLevel::LOG_DEBUG, 
          std::move(xlo_msg));
      }

      XLO_DEBUG("{0} (in {2}:{3} during {1})", msg, func, filename, line);
    }
    catch (...)
    {
    }
  }

  std::wstring writeWindowsError()
  {
    return Helpers::writeWindowsError();
  }
}