#pragma once
#include <memory>
namespace Excel { struct AppEvents; struct _Application; }
namespace xloil
{
  namespace COM
  {
    std::shared_ptr<Excel::AppEvents> createEventSink(Excel::_Application* source);
  }
}