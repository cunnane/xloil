#pragma once

namespace xloil
{
  /// Load plugins according to settings file. Called by xlAutoOpen
  void loadPlugins() noexcept;

  /// Unloads any plugins prior to takedown of the Core XLL. 
  /// Called by xlAutoClose
  void unloadPlugins() noexcept;
}
