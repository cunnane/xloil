#pragma once
#include <xlOil/Interface.h>
#include <xlOil/Log.h>

namespace xloil
{
  /// <summary>
  /// Links a plug-in's *spdlog* instance to the main xlOil log output. 
  /// You don't have to do this if you're organising your own logging.
  /// </summary>
  /// <param name=""></param>
  /// <param name="plugin"></param>
  inline void linkLogger(AddinContext*, const PluginContext& plugin)
  {
    if (plugin.action == PluginContext::Load)
      spdlog::set_default_logger(loggerRegistry().default_logger());
  }
}