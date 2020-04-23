#include <xlOil/StaticRegister.h>
#include <xloil/Interface.h>
#include "Cache.h"
#include <xlOil/Log.h>

namespace xloil
{
  namespace SQL
  {
    XLO_PLUGIN_INIT(AddinContext* context, const PluginContext& plugin)
    {
      if (plugin.action == PluginContext::Load)
      {
        spdlog::set_default_logger(context->getLogger());
        createCache();
      }
      return 0;
    }
  }
}


