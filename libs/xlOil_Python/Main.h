#pragma once
#include <xloil/Interface.h>

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
  }
}