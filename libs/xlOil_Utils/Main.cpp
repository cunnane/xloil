#include <xloil/Interface.h>

namespace xloil
{
  namespace Utils
  {
    Core* theCore = nullptr;

    XLO_PLUGIN_INIT(Core& core)
    {
      theCore = &core;
      spdlog::set_default_logger(core.getLogger());
      return 0;
    }
  }
}

