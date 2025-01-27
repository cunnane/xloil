#pragma once
#include <memory>
#include <set>
#include <string>
namespace Excel { struct AppEvents; struct _Application; }
namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Creating the event sink may trigger events which may in turn rely on the COM 
    /// connection being available - ensure that it is before calling this function.
    /// </summary>
    std::shared_ptr<Excel::AppEvents> createEventSink(Excel::_Application* source);

    const std::set<std::wstring>& workbookPaths();
  }
}