#include "Events.h"
#include <xlOil/Log.h>
#include <map>
#include <vector>
#include <unordered_map>
#include <array>
#include <simplefilewatcher/include/FileWatcher/FileWatcher.h>
#include <string>

using std::vector;
using std::shared_ptr;
using std::make_shared;
using std::unordered_map;
using std::wstring;

namespace xloil
{
  namespace Event
  {
    using namespace detail;
    
    // Not exported, so separate
    EventXll& AutoClose()
    { 
      static EventXll e("AutoClose"); return e;
    }

#define XLO_DEF_EVENT(name) \
    XLOIL_EXPORT decltype(name()) name() \
    { \
      static std::remove_reference_t<decltype(name())> e(#name); \
      return e; \
    }

    XLO_DEF_EVENT(AfterCalculate);
    XLO_DEF_EVENT(CalcCancelled);
    XLO_DEF_EVENT(WorkbookOpen);
    XLO_DEF_EVENT(NewWorkbook);
    XLO_DEF_EVENT(SheetSelectionChange);
    XLO_DEF_EVENT(SheetBeforeDoubleClick);
    XLO_DEF_EVENT(SheetBeforeRightClick);
    XLO_DEF_EVENT(SheetActivate);
    XLO_DEF_EVENT(SheetDeactivate);
    XLO_DEF_EVENT(SheetCalculate);
    XLO_DEF_EVENT(SheetChange);
    XLO_DEF_EVENT(WorkbookAfterClose);
    XLO_DEF_EVENT(WorkbookActivate);
    XLO_DEF_EVENT(WorkbookDeactivate);
    XLO_DEF_EVENT(WorkbookBeforeSave);
    XLO_DEF_EVENT(WorkbookBeforePrint);
    XLO_DEF_EVENT(WorkbookNewSheet);
    XLO_DEF_EVENT(WorkbookAddinInstall);
    XLO_DEF_EVENT(WorkbookAddinUninstall);

    using DirectoryWatchEvent = Event<void(const wchar_t*, const wchar_t*, FileAction)>;

    static FW::FileWatcher theFileWatcher;

    class DirectoryListener : public FW::FileWatchListener
    {
    public:
      DirectoryListener(const std::wstring& path, std::function<void(void)> finaliser)
        : _eventSource()
        , _lastTickCount(0)
        , _watchId(theFileWatcher.addWatch(path, this, false))
      {
      }

      ~DirectoryListener()
      {
        theFileWatcher.removeWatch(_watchId);
      }

      void handleFileAction(FW::WatchID, const std::wstring& dir, const std::wstring& filename,
        FW::Action action) override
      {
        // File updates seem to generate two identical calls so implement a time granularity
        auto ticks = GetTickCount64();
        if (ticks - _lastTickCount < 1000)
          return;
        _lastTickCount = ticks;
        _eventSource.fire(dir.c_str(), filename.c_str(), FileAction(action));
      }

      DirectoryWatchEvent& event() { return _eventSource; }

    private:
      FW::WatchID _watchId;
      DirectoryWatchEvent _eventSource;
      size_t _lastTickCount;
    };

    static unordered_map<wstring, shared_ptr<DirectoryListener>> theDirectoryListeners;

    XLOIL_EXPORT DirectoryWatchEvent& DirectoryChange(const std::wstring& path)
    {
      auto found = theDirectoryListeners.find(path);
      if (found != theDirectoryListeners.end())
        return found->second->event();

      wstring pathStr(path);
      auto[it, ins] = theDirectoryListeners.emplace(
        pathStr, new DirectoryListener(pathStr, [pathStr]() { theDirectoryListeners.erase(pathStr); }));
      return it->second->event();
    }
  }
}