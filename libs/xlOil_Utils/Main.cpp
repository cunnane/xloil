#include <xlOil/Plugin.h>
#include <xloil/ExcelObj.h>

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

XLO_DEFINE_FREE_CALLBACK()
