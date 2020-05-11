#include "Events.h"
#include <xlOil/Log.h>
#include <map>
#include <vector>
#include <unordered_map>
#include <array>
#include <simplefilewatcher/include/FileWatcher/FileWatcher.h>
#include <string>
#include <boost/preprocessor/seq/for_each.hpp>

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


#define XLO_DEF_EVENT(r, _, name) \
    XLOIL_EXPORT decltype(name()) name() \
    { \
      static std::remove_reference_t<decltype(name())> e(#name); \
      return e; \
    };

    BOOST_PP_SEQ_FOR_EACH(XLO_DEF_EVENT, _, XLOIL_STATIC_EVENTS)

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