#include <xlOil/Plugin.h>

namespace xloil
{
  namespace Utils
  {
    XLO_PLUGIN_INIT(AddinContext* ctx, const PluginContext& plugin)
    {
      try
      {
        linkPluginToCoreLogger(ctx, plugin);
        throwIfNotExactVersion(plugin);
        return 0;
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(e.what());
        return -1;
      }
    }
  }
}

