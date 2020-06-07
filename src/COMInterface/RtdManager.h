#pragma once
#include <xloil/RtdServer.h>

namespace xloil
{
  namespace COM
  {
    std::shared_ptr<IRtdManager> newRtdManager(
      const wchar_t* progId, const wchar_t* clsid);
  }
}
