#include <xlOil/Plugin.h>

namespace xloil
{
  namespace Utils
  {
    XLO_PLUGIN_INIT(AddinContext* ctx, const PluginContext& plugin)
    {
      linkPluginToCoreLogger(ctx, plugin);
      return 0;
    }
  }
}

