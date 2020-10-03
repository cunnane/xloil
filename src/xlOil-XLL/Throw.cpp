#include <xloil/Throw.h>
#include <xloil/Log.h>
#include <xlOilHelpers/Exception.h>

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
    {
    }
  }


  std::wstring writeWindowsError()
  {
    return Helpers::writeWindowsError();
  }
}