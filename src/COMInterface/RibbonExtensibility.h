#pragma once
#include <map>
#include <functional>

namespace Office { struct IRibbonExtensibility; }
namespace xloil { struct RibbonControl;  }

namespace xloil
{
  namespace COM
  {
    Office::IRibbonExtensibility* createRibbon(
      const wchar_t* xml,
      const std::map<std::wstring, std::function<void(const RibbonControl&)>> handlers);
  }
}