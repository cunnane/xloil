#pragma once
#include <vector>
#include <string>

namespace xloil
{
  class AddinContext;

  /// Load plugins according to settings file. Called by xlAutoOpen
  void loadPlugins(AddinContext* context, const std::vector<std::wstring>& names) noexcept;

  /// Unloads any plugins prior to takedown of the Core XLL. 
  /// Called by xlAutoClose
  void unloadPlugins() noexcept;
}
