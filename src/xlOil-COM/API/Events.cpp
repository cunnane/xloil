#include <xlOil/Events.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/StringUtils.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil-COM/Connect.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/State.h>
#include <xlOil/Throw.h>
#include <map>
#include <vector>
#include <unordered_map>
#include <array>
#include <rdcfswatcher/rdc_fs_watcher.h>
#include <string>
#include <boost/preprocessor/seq/for_each.hpp>
#include <boost/preprocessor/stringize.hpp>

using std::vector;
using std::shared_ptr;
using std::weak_ptr;
using std::make_shared;
using std::unordered_map;
using std::wstring;
using std::set;
using std::wstring;
using std::pair;

namespace xloil
{
  namespace Event
  {
    namespace
    {
      void directoryWatchChange(int64_t id, const set<pair<wstring, uint32_t>>& notifications);

      void directoryWatchError(int64_t id);

      auto& theFileWatcher()
      {
        static std::unique_ptr<RdcFSWatcher> instance([]()
        {
          auto instance = new RdcFSWatcher();
          instance->changeEvent = directoryWatchChange;
          instance->errorEvent = directoryWatchError;
          return instance;
        }());
        return *instance;
      }

      using DirectoryWatchEventBase = Event<void(const wchar_t*, const wchar_t*, FileAction)>;

      struct DirectoryWatchEvent : public DirectoryWatchEventBase
      {
        DirectoryWatchEvent(const std::wstring& path)
          : DirectoryWatchEventBase((L"Watch_" + path).c_str())
          , _lastTickCount(0)
          , _directory(path)
        {
          theFileWatcher().addDirectory((intptr_t)this, path);
          XLO_DEBUG(L"Started directory watch on '{}'", path);
        }

        virtual ~DirectoryWatchEvent()
        {
          XLO_DEBUG(L"Ended directory watch '{}'", name());
          theFileWatcher().removeDirectory((intptr_t)this);
        }

        void handleFileAction(
          const wstring& filename,
          FileAction action)
        {
          if (filename.find(L'\\') != wstring::npos)
            return;

          // File updates seem to generate two identical calls so implement a time granularity
          auto ticks = GetTickCount64();
          if (ticks - _lastTickCount < 1000)
            return;
          _lastTickCount = ticks;

          this->fire(_directory.c_str(), filename.c_str(), action);
        }

        auto& directory() const { return _directory; }

      private:
        decltype(GetTickCount64()) _lastTickCount;
        wstring _directory;
      };

      void directoryWatchChange(int64_t id, const set<pair<wstring, uint32_t>>& notifications)
      {
        auto target = (DirectoryWatchEvent*)id;
        for (const auto& notification : notifications)
        {
          FileAction fwAction;
          switch (notification.second)
          {
          case FILE_ACTION_RENAMED_NEW_NAME:
          case FILE_ACTION_ADDED:
            fwAction = FileAction::Add;
            break;
          case FILE_ACTION_RENAMED_OLD_NAME:
          case FILE_ACTION_REMOVED:
            fwAction = FileAction::Delete;
            break;
          case FILE_ACTION_MODIFIED:
            fwAction = FileAction::Modified;
            break;
          default:
            return;
          };
          target->handleFileAction(notification.first, fwAction);
        }
      }

      void directoryWatchError(int64_t id)
      {
        auto target = (DirectoryWatchEvent*)id;
        XLO_ERROR(L"Directory watcher for '{}' failed: {}", target->directory(), writeWindowsError());
      };

      unordered_map<wstring, weak_ptr<DirectoryWatchEvent>> theDirectoryWatchers;
    }

    XLOIL_EXPORT std::wstring to_wstring(const FileAction x)
    {
      switch (x)
      {
      case FileAction::Add: return L"add";
      case FileAction::Delete: return L"delete";
      case FileAction::Modified: return L"modified";
      default:
        return L"unknown";
      }
    }

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

    // This event is not exported, so defined separately
    XLOIL_DEFINE_EVENT(AutoClose);

    // Define all standard events
#define XLO_DEF_EVENT(r, _, name) XLOIL_EXPORT XLOIL_DEFINE_EVENT(name);
    BOOST_PP_SEQ_FOR_EACH(XLO_DEF_EVENT, _, XLOIL_STATIC_EVENTS)
#undef XLO_DEF_EVENT
  }
}