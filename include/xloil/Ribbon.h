#pragma once
#include "ExportMacro.h"
#include <functional>
#include <memory>
#include <map>

namespace xloil
{
  struct RibbonControl
  {
    const wchar_t* Id;
    const wchar_t* Tag;
  };

  class IComAddin
  {
  public:
    using Handlers = std::map<std::wstring, std::function<void(const RibbonControl&)>>;
    virtual ~IComAddin() {}
    virtual const wchar_t* progid() const = 0;
    virtual void connect() = 0;
    virtual void disconnect() = 0;
    virtual void setRibbon(
      const wchar_t* xml,
      const Handlers& handlers) = 0;
    virtual void ribbonInvalidate(const wchar_t* controlId = 0) const = 0;
    virtual bool ribbonActivate(const wchar_t* controlId) const = 0;
  };

  XLOIL_EXPORT std::shared_ptr<IComAddin>
    makeComAddin(const wchar_t* name, const wchar_t* description = nullptr);
}