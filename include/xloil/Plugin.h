#pragma once
#include <xlOil/Interface.h>
#include <xlOil/Log.h>
#include <xlOil/Throw.h>

namespace xloil
{
  /// <summary>
  /// Throws if the version of xlOil used to build the plugin does not exactly match
  /// the one loading it. Plugins could apply more sophisticated logic based on API 
  /// changes, but calling this is a safe default.
  /// </summary>
  /// <param name="ctx"></param>
  inline void throwIfNotExactVersion(const PluginContext& ctx)
  {
    if (ctx.action == PluginContext::Load && !ctx.checkExactVersion())
      XLO_THROW(L"Plugin '{}' expected xlOil %d.%d.%d but was loaded by %d.%d.%d",
        ctx.pluginName, XLOIL_MAJOR_VERSION, XLOIL_MINOR_VERSION, XLOIL_PATCH_VERSION,
        ctx.coreMajorVersion, ctx.coreMinorVersion, ctx.corePatchVersion);
  }
}