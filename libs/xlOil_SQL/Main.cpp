#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include "Cache.h"


namespace xloil
{
  namespace SQL
  {
    Core* theCore = nullptr;

    XLO_PLUGIN_INIT(xloil::Core& core)
    {
      theCore = &core;
      spdlog::set_default_logger(core.getLogger());
      createCache();
      return 0;
    }
  }
}


