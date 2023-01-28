#pragma once
#include <memory>
#include <set>
#include <string>
namespace Excel { struct AppEvents; struct _Application; }
namespace xloil
{
  namespace COM
  {
    std::shared_ptr<Excel::AppEvents> createEventSink(Excel::_Application* source);
    const std::set<std::wstring>& workbookPaths();
  }
}