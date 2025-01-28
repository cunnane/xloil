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
using std::shared_mutex;


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

    using CellAddress = std::pair<unsigned, unsigned>;

    using CellTaskMap = std::unordered_map<
      CellAddress,
      std::shared_ptr<CellTasks>,
      pair_hash<unsigned, unsigned>>;

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
        if (numSubscribers != 0)
          XLO_ERROR(L"AsyncTaskPublisher: unexpected subscribers for {}", topic());
        RtdPublisher::disconnect(numSubscribers);
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
      CellTaskMap& tasksPerCell,
      const CellAddress& address,
      const msxll::XLREF12& ref)
    {
      auto [iTask, success] = tasksPerCell.try_emplace(address, new CellTasks());
      auto tasksInCell = iTask->second;
      tasksInCell->setCaller(ref);
      return tasksInCell;
    }

    void writeArray(
      CellTaskMap& tasksPerCell,
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


    namespace RtdAsyncManager
    {
      namespace
      {
        // Hold this mutex to a create the Impl class or to call its methods
        static shared_mutex theManagerMutex;

        class Impl
        {
        public:
          Impl()
            : _rtd(newRtdServer(nullptr, nullptr))
          {
            // We're a singleton so guaranteed to still exist at autoclose
            Event::AutoClose() += [this]() {
              clear();
              _rtd.reset();
            };
          }

          void clear()
          {
            _rtd->clear();
            _tasksPerCell.clear();
          }

          auto findTargetCellTasks(
            const msxll::XLREF12* ref,
            unsigned arraySize,
            unsigned sheetId,
            unsigned cellNumber)
          {
            // Lock the dictionary of cell tasks and look for the address
            // (1) New master (no previous record)
            // (2) New master (former slave)
            //     (a) Shares top left
            //     (b) Doesn't
            // (3) Slave increment (arraySize = 1, arrayCount > 0)

            // The value we need to populate
            shared_ptr<CellTasks> pTasksInCell;

            const auto address = std::make_pair(sheetId, cellNumber);

            const auto found = _tasksPerCell.find(address);
            if (found == _tasksPerCell.end())
            {
              pTasksInCell = newCellTasks(_tasksPerCell, address, *ref);
            }
            else
            {
              pTasksInCell = found->second;
              auto tasksInCell = pTasksInCell.get();

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

            return pTasksInCell;
          }
          const std::shared_ptr<IRtdServer>& getRtd() const
          {
            return _rtd;
          }
         
        private:
          std::shared_ptr<IRtdServer> _rtd;
          CellTaskMap _tasksPerCell;
        };

        static std::unique_ptr<Impl> theInstance;

        Impl* getInstance(unique_lock<shared_mutex>& lock)
        {
          // Somewhat intricate logic around creating the instance:
          //  * If we are the main thread, just create it
          //  * Otherwise release the lock and send a request to 
          //    main thread to acquire the lock and create it.
          //  * Wait for 5s, then give up
          //  * We need the lock to call functions on theInstance
          //    so reacquire it if lost
          if (!theInstance)
          {
            if (isMainThread())
              theInstance.reset(new Impl());
            else
            {
              lock.unlock();

              auto status = runExcelThread([]
              {
                std::unique_lock lock(theManagerMutex);
                if (!theInstance)
                  theInstance.reset(new Impl());
              }).wait_for(std::chrono::seconds(5));

              if (status != std::future_status::ready)
                XLO_THROW("Rtd async manager timed out, try calling the function again");

              // Instance now exists, so get lock and get cracking
              lock.lock();
            }
          }
          return theInstance.get();
        }

        /// <summary>
        /// Static, so does not need the mutex
        /// </summary>
        static auto getValueAndSubscribe(
          const std::shared_ptr<IRtdServer>& rtd,
          const std::shared_ptr<IRtdAsyncTask>& task,
          const unsigned arraySize,
          const shared_ptr<CellTasks>& pTasksInCell)
        {
          auto* tasksInCell = pTasksInCell.get();

          // Now populate these variables
          shared_ptr<const ExcelObj> result;
          const wchar_t* foundTopic = nullptr;

          {
            // Lock 'tasksInCell' in case there is more than one RTD function in the cell.
            // This is unlikely in itself and as they all execute on the same thread, contention
            // problems are improbable, so we use a lightweight atomic_flag which implies a spin 
            // wait (before C++20).
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
                    result = rtd->peek(foundTopic);
                  break;
                }

              if (!foundTopic)
              {
                // Couldn't find a matching task so start a new one
                startCellTask(*rtd, pTasksInCell, task);
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
          if (result)
            return result;
       
          // If the task is still running we resubscribe to all other tasks in the cell.
          // If an argument to this task is the result of another task in the same cell,
          // not subscribing will cause Excel to send a disconnect to the inner task. When
          // this task returns its result, it triggers a cell recalc which will start a new
          // connection to the inner task and go into a perpetual loop. If the inner task
          // is still connected, xlOil knows it has already produced a result and returns 
          // it, avoiding the loop.
          for (auto& t : tasksInCell->tasks)
            if (wcscmp(t->topic(), foundTopic) != 0)
              rtd->subscribe(t->topic());
          return rtd->subscribe(foundTopic);
        }
      }

      void init()
      {
        // Acquire the lock then check if the instance has been created
        unique_lock lock(theManagerMutex);
        getInstance(lock);
      }

      std::shared_ptr<const ExcelObj>
        getValue(
          const std::shared_ptr<IRtdAsyncTask>& task)
      {
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
        const auto sheetId = (unsigned)(intptr_t)callExcel(
          msxll::xlSheetId, caller.fullSheetName()).val.mref.idSheet;

        // Acquire the lock then check if the instance has been created
        unique_lock lock(theManagerMutex);

        auto impl = getInstance(lock);

        auto tasksInCell = impl->findTargetCellTasks(
          ref, arraySize, cellNumber, sheetId);

        auto rtdServer = impl->getRtd();

        // We've finished with the task-per-cell lookup and can drop the lock
        // required for the Impl
        lock.unlock();

        return getValueAndSubscribe(rtdServer, task, arraySize, tasksInCell);
      }

      void clear()
      {
        unique_lock lock(theManagerMutex);

        if (!theInstance)
          return;

        theInstance->clear();
      }
    } // namespace RtdAsyncManager
  }
}