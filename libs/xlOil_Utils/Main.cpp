#include <xloil/Interface.h>
#include <xlOil/Log.h>

namespace xloil
{
  namespace Utils
  {
    XLO_PLUGIN_INIT(AddinContext* ctx, const PluginContext& plugin)
    {
      if (plugin.action == PluginContext::Load)
        spdlog::set_default_logger(ctx->getLogger());
      return 0;
    }
  }
}

