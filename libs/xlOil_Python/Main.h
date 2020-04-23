#pragma once
#include <xloil/Interface.h>
#include <xlOil/Events.h>

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// The addin context of the main xloil.dll
    /// </summary>
    extern AddinContext* theCoreContext;

    /// <summary>
    /// The current context is set to reflect the addin whose
    /// settings are being processed. It is then switched back
    /// to the core context.
    /// </summary>
    extern AddinContext* theCurrentContext;

    /// <summary>
    /// An event triggered when the Python plugin is about to close
    /// but before the Python interpreter is stopped.
    /// </summary>
    Event<void(void), VoidCollector>& Event_PyBye();
  }
}