#pragma once
#include <xloil/RtdServer.h>

namespace xloil
{
  namespace COM
  {
    std::shared_ptr<IRtdServer> newRtdServer(
      const wchar_t* progId, const wchar_t* clsid);
  }
}
