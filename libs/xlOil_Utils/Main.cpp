#include <xloil/Interface.h>
#include <xlOil/Log.h>

namespace xloil
{
  namespace Utils
  {
    XLO_PLUGIN_INIT(AddinContext* ctx, const PluginContext& plugin)
    {
      linkLogger(ctx, plugin);
      return 0;
    }
  }
}

