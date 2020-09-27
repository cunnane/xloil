#pragma once
#include <map>
#include <functional>
#include <memory>

namespace Office { struct IRibbonExtensibility; }
namespace xloil { struct RibbonControl;  }

namespace xloil
{
  namespace COM
  {
    class IRibbon
    {
    public:
      virtual Office::IRibbonExtensibility* getRibbon() = 0;
      virtual void invalidate(const wchar_t* controlId = 0) const = 0;
      virtual bool activateTab(const wchar_t* controlId) const = 0;
    };
    std::shared_ptr<IRibbon> createRibbon(
      const wchar_t* xml,
      const std::map<std::wstring, std::function<void(const RibbonControl&)>>& handlers);
  }
}