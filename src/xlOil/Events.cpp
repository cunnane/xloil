#include "Events.h"
#include "Log.h"
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
  XLOIL_EXPORT Event<void(void), VoidCollector>& Event_AutoOpen() { static Event<void(void), VoidCollector> e; return e; }
  Event<void(void), VoidCollector>& Event_AutoClose() { static Event<void(void), VoidCollector> e; return e; }
  XLOIL_EXPORT Event<void(void), VoidCollector>& Event_CalcEnded() { static Event<void(void), VoidCollector> e; return e; }
  XLOIL_EXPORT Event<void(void), VoidCollector>& Event_CalcCancelled() { static Event<void(void), VoidCollector> e; return e; }
  XLOIL_EXPORT Event<void(const wchar_t*, const wchar_t*), VoidCollector>& Event_WorkbookOpen() { static Event<void(const wchar_t*, const wchar_t*), VoidCollector> e; return e; }
  XLOIL_EXPORT Event<void(const wchar_t*), VoidCollector>& Event_NewWorkbook() { static Event<void(const wchar_t*), VoidCollector> e; return e; }
  XLOIL_EXPORT Event<void(const wchar_t*), VoidCollector>& Event_WorkbookClose() { static Event<void(const wchar_t*), VoidCollector> e; return e; }

  using DirectoryWatchEvent = Event<void(const wchar_t*, const wchar_t*, FileAction), VoidCollector>;

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

    void handleFileAction(FW::WatchID watchid, const std::wstring& dir, const std::wstring& filename,
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

  XLOIL_EXPORT DirectoryWatchEvent& Event_DirectoryChange(const std::wstring& path)
  {
    auto found = theDirectoryListeners.find(path);
    if (found != theDirectoryListeners.end())
      return found->second->event();

    wstring pathStr(path);
    auto[it, ins] = theDirectoryListeners.emplace(
      pathStr, new DirectoryListener(pathStr, [pathStr](){ theDirectoryListeners.erase(pathStr); }));
    return it->second->event();
  }
}