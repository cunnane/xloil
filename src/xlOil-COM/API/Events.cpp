#include <xlOil/Events.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/StringUtils.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil-COM/Connect.h>
#include <map>
#include <vector>
#include <unordered_map>
#include <array>
#include <simplefilewatcher/include/FileWatcher/FileWatcher.h>
#include <string>
#include <boost/preprocessor/seq/for_each.hpp>
#include <boost/preprocessor/stringize.hpp>

using std::vector;
using std::shared_ptr;
using std::weak_ptr;
using std::make_shared;
using std::unordered_map;
using std::wstring;

namespace xloil
{
  namespace Event
  {
    using namespace detail;
    
    static FW::FileWatcher theFileWatcher;

    // Not exported, so separate
    XLOIL_DEFINE_EVENT(AutoClose)

#define XLO_DEF_EVENT(r, _, name) XLOIL_EXPORT XLOIL_DEFINE_EVENT(name)
      BOOST_PP_SEQ_FOR_EACH(XLO_DEF_EVENT, _, XLOIL_STATIC_EVENTS)
#undef XLO_DEF_EVENT

    using DirectoryWatchEventBase = Event<void(const wchar_t*, const wchar_t*, FileAction)>;

    struct DirectoryWatchEvent : public DirectoryWatchEventBase, public FW::FileWatchListener
    {
      DirectoryWatchEvent(const std::wstring& path)
        : DirectoryWatchEventBase(("Watch_" + utf16ToUtf8(path)).c_str())
        , _lastTickCount(0)
        , _watchId(theFileWatcher.addWatch(path, this, false))
      {
        XLO_TRACE(L"Started directory watch on '{}'", path);
      }

      virtual ~DirectoryWatchEvent()
      {
        XLO_TRACE("Ended directory watch '{}'", name());
        theFileWatcher.removeWatch(_watchId);
      }

      void handleFileAction(
        FW::WatchID,
        const std::wstring& dir,
        const std::wstring& filename,
        FW::Action action) override
      {
        // File updates seem to generate two identical calls so implement a time granularity
        auto ticks = GetTickCount64();
        if (ticks - _lastTickCount < 1000)
          return;
        _lastTickCount = ticks;
        this->fire(dir.c_str(), filename.c_str(), FileAction(action));
      }
    private:
      FW::WatchID _watchId;
      size_t _lastTickCount;
    };


    static unordered_map<wstring, weak_ptr<DirectoryWatchEvent>> theDirectoryWatchers;

    XLOIL_EXPORT shared_ptr<DirectoryWatchEventBase> DirectoryChange(const std::wstring& path)
    {
      auto found = theDirectoryWatchers.find(path);
      if (found != theDirectoryWatchers.end())
      {
        auto ptr = found->second.lock();
        // If our weak_ptr is dead, erase it or the emplace below will fail
        if (ptr)
          return ptr;
        else 
          theDirectoryWatchers.erase(found);
      }

      auto event = make_shared<DirectoryWatchEvent>(path);
      auto [it, ins] = theDirectoryWatchers.emplace(path, event);
      return event;
    }

    XLOIL_EXPORT void allowEvents(bool value)
    {
      try
      {
        COM::excelApp().EnableEvents = _variant_t(value);
      }
      XLO_RETHROW_COM_ERROR;
    }
  }
}