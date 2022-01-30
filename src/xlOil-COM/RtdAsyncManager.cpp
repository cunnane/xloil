#include "RtdAsyncManager.h"
#include "RtdManager.h"
#include <xlOil/RtdServer.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Caller.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelThread.h>
#include <xloil/StringUtils.h>
#include <combaseapi.h>
#include <shared_mutex>

using std::wstring;
using std::shared_ptr;
using std::unique_ptr;
using std::make_shared;
using std::shared_lock;
using std::make_pair;
using std::unique_lock;


namespace
{
  /// <summary>
  /// Like a std::scoped_lock but uses a std::atomic_flag rather than a mutex.
  /// Note it busy-waits for the lock!
  /// </summary>
  struct scoped_atomic_flag
  {
    std::atomic_flag* _flag;

    scoped_atomic_flag(std::atomic_flag& flag)
      : _flag(&flag)
    {
      while (flag.test_and_set(std::memory_order_acquire))
      {
        // Since C++20, it is possible to update atomic_flag's
        // value only when there is a chance to acquire the lock.
        // See also: https://stackoverflow.com/questions/62318642
#if defined(__cpp_lib_atomic_flag_test)
        while (lock.test(std::memory_order_relaxed))
#else
        // spin
#endif
      }
    }
    ~scoped_atomic_flag()
    {
      if (_flag)
        _flag->clear(std::memory_order_release);
    }
    void swap(scoped_atomic_flag& that)
    {
      std::swap(_flag, that._flag);
    }
  };
}

namespace xloil
{
  namespace COM
  {
    class AsyncTaskPublisher;

    struct CellTasks
    {
      std::list<shared_ptr<AsyncTaskPublisher>> tasks;
      int arrayCount = 0; // see comment in 'getValue()'
      const wchar_t* arrayTopic = nullptr;
      msxll::XLREF12 caller;
      std::atomic_flag busy = ATOMIC_FLAG_INIT;

      bool isSubarray(const msxll::XLREF12& ref) const
      {
        return ref.rwFirst >= caller.rwFirst && ref.rwLast <= caller.rwLast
          && ref.colFirst >= caller.colFirst && ref.colLast <= caller.colLast;
      }
      void setCaller(const msxll::XLREF12& ref)
      {
        memcpy(&caller, &ref, sizeof(msxll::XLREF12));
      }
    };

    // TODO: could we just create a forwarding IRtdAsyncTask which intercepts 'cancel'
    class AsyncTaskPublisher : public RtdPublisher
    {
      std::weak_ptr<CellTasks> _parent;

    public:
      AsyncTaskPublisher(
        const wchar_t* topic,
        IRtdServer& mgr,
        const shared_ptr<IRtdTask>& task,
        const std::weak_ptr<CellTasks>& parent)
        : RtdPublisher(topic, mgr, task)
        , _parent(parent)
      {}

      bool disconnect(size_t numSubscribers) override
      {
        RtdPublisher::disconnect(numSubscribers);
        // TODO: check numSubscribers == 0
        stop();
        auto p = _parent.lock();
        if (p)
        {
          scoped_atomic_flag lock(p->busy);
          p->tasks.remove_if([&](auto& t) { return t.get() == this; });
        }
        return true;
      }

      bool taskMatches(const IRtdAsyncTask& that) const
      {
        return that == (const IRtdAsyncTask&)*task();
      }
    };

    void startCellTask(
      IRtdServer& _rtd,
      const shared_ptr<CellTasks>& tasks, 
      const shared_ptr<IRtdAsyncTask>& task)
    {
      GUID guid;
      wchar_t guidStr[64];

      if (CoCreateGuid(&guid) != 0 || StringFromGUID2(guid, guidStr, _countof(guidStr)) == 0)
        XLO_THROW("Internal: RtdAsyncManager failed to create GUID");

      tasks->tasks.emplace_back(new AsyncTaskPublisher(guidStr, _rtd, task, tasks));

      _rtd.start(tasks->tasks.back());
    }

    auto newCellTasks(
      RtdAsyncManager::CellTaskMap& tasksPerCell,
      const RtdAsyncManager::CellAddress& address,
      const msxll::XLREF12& ref)
    {
      auto [iTask, success] = tasksPerCell.try_emplace(address, new CellTasks());
      auto tasksInCell = iTask->second;
      tasksInCell->setCaller(ref);
      return tasksInCell;
    }

    void writeArray(
      RtdAsyncManager::CellTaskMap& tasksPerCell,
      const shared_ptr<CellTasks>& val,
      const unsigned sheetId,
      const msxll::XLREF12& ref)
    {
      for (auto j = ref.colFirst; j <= ref.colLast; ++j)
        for (auto i = ref.rwFirst; i <= ref.rwLast; ++i)
        {
          unsigned num = i * XL_MAX_COLS + j;
          tasksPerCell[make_pair(sheetId, num)] = val;
        }
    }

      RtdAsyncManager::RtdAsyncManager() 
        : _rtd(newRtdServer(nullptr, nullptr))
      {
        // We're a singleton so guaranteed to still exist at autoclose
        Event::AutoClose() += [this]() {
          clear();
          _rtd.reset();
        };
      }

      RtdAsyncManager& RtdAsyncManager::instance()
      {
        static RtdAsyncManager* mgr = runExcelThread([]()
        {
          return new RtdAsyncManager();
        }).get();
        return *mgr;
      }

      shared_ptr<const ExcelObj> RtdAsyncManager::getValue(
        const std::shared_ptr<IRtdAsyncTask>& task)
      {
        // Protects agains a null-deref and allows starting up the RTD server
        // without running anything
        if (!task)
          return shared_ptr<const ExcelObj>();

        const auto caller = CallerInfo();
        const auto ref = caller.sheetRef();

        // Sometimes fetching the caller fails - give up 
        if (!ref)
          return shared_ptr<const ExcelObj>();

        const auto arraySize = (ref->colLast - ref->colFirst + 1)
          * (ref->rwLast - ref->rwFirst + 1);

        // This is the cell number of the top-left cell for array callers
        const unsigned cellNumber = ref->rwFirst * XL_MAX_COLS + ref->colFirst;

        // It's OK to cast away half the sheetref ptr: it's likely the detail
        // is in the lower part and it doesn't matter if we have collisions
        // in the map since we check for explicit equality later.
        const auto sheetId = (unsigned)(intptr_t)callExcel(msxll::xlSheetId, caller.fullSheetName()).val.mref.idSheet;
        const auto address = std::make_pair(sheetId, cellNumber);

        // The values we need to populate
        shared_ptr<CellTasks> pTasksInCell;
        CellTasks* tasksInCell;

        // Lock the dictionary of cell tasks and look for the address
        // (1) New master (no previous record)
        // (2) New master (former slave)
        //     (a) Shares top left
        //     (b) Doesn't
        // (3) Slave increment (arraySize = 1, arrayCount > 0)

        unique_lock writeLock(_mutex);
        const auto found = _tasksPerCell.find(address);
        if (found == _tasksPerCell.end())
        {
          pTasksInCell = newCellTasks(_tasksPerCell, address, *ref);
          tasksInCell = pTasksInCell.get();
        }
        else
        {
          pTasksInCell = found->second;
          tasksInCell = pTasksInCell.get();

          if (!pTasksInCell->isSubarray(*ref))
          {
            pTasksInCell = newCellTasks(_tasksPerCell, address, *ref);
          }
          else if (arraySize == 1 && tasksInCell->arrayCount > 0)
          {
            // Do nothing for now
          }
          else
          {
            tasksInCell->setCaller(*ref);
          }
        }

        // If the caller is an array formula, when RTD is called in the subscribe()
        // method, it will return xlretUncalced, but will trigger the calling
        // function to be called again for each cell in the array. The caller in these
        // subsequent calls will sometimes remain as the top left-cell and will sometimes 
        // cycle through the cells of the array.
        // 
        // We want to start the task only once for the first call, with subsequent
        // calls invoking subscribe quickly without needing to compare all function args.

        if (arraySize > 1)
          writeArray(_tasksPerCell, pTasksInCell, sheetId, *ref);

        // We've finished with the task-per-cell lookup
        writeLock.unlock();

        // Now populate these variables
        shared_ptr<const ExcelObj> result;
        const wchar_t* foundTopic = nullptr;
        {
          // Lock 'tasksInCell' in case there is more than one RTD function in the cell
          // This is unlikely, so we use a lightweight atomic_flag which implies a spin wait
          // (before C++20).
          scoped_atomic_flag lockCell(tasksInCell->busy);

          if (arraySize == 1 && tasksInCell->arrayCount > 0)
          {
            --tasksInCell->arrayCount;
            foundTopic = tasksInCell->arrayTopic;
          }
          else
          {
            // Compare our task to all other running tasks in the cell to see if we
            // already have the answer
            for (auto& t : tasksInCell->tasks)
              if (t->taskMatches(*task))
              {
                foundTopic = t->topic();
                if (t->done())
                  result = _rtd->peek(foundTopic);
                break;
              }

            if (!foundTopic)
            {
              // Couldn't find a matching task so start a new one
              startCellTask(*_rtd, pTasksInCell, task);
              foundTopic = tasksInCell->tasks.back()->topic();
            }

            if (arraySize > 1)
            {
              tasksInCell->arrayCount = arraySize;
              tasksInCell->arrayTopic = foundTopic;
            }
            else
            {
              tasksInCell->arrayCount = 0;
              tasksInCell->arrayTopic = nullptr;
            }
          }
        }
        assert(foundTopic);
        return result ? result : _rtd->subscribe(foundTopic);
      }

      void RtdAsyncManager::clear()
      {
        std::unique_lock lock(_mutex);
        _rtd->clear();
        _tasksPerCell.clear();
      }
    };
  }