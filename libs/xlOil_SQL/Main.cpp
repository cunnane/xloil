#include <xlOil/StaticRegister.h>
#include <xloil/Plugin.h>
#include "Cache.h"

namespace xloil
{
  namespace SQL
  {
    XLO_PLUGIN_INIT(AddinContext* context, const PluginContext& plugin)
    {
      throwIfNotExactVersion(plugin);
      linkPluginToCoreLogger(context, plugin);
      return 0;
    }
  }
}


