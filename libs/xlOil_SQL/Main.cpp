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
      linkLogger(context, plugin);

      if (plugin.action == PluginContext::Load)
        createCache();
      
      return 0;
    }
  }
}


