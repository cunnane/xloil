#include <xlOil/StaticRegister.h>
#include <xloil/Plugin.h>
#include "Cache.h"

namespace xloil
{
  namespace SQL
  {
    XLO_PLUGIN_INIT(AddinContext* context, const PluginContext& plugin)
    {
      try
      {
        linkPluginToCoreLogger(context, plugin);
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
