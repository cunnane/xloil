#include <xlOil/Log.h>

namespace xloil
{
  XLOIL_EXPORT spdlog::details::registry& loggerRegistry()
  {
    return spdlog::details::registry::instance();
  }
}