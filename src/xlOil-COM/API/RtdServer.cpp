#include <xlOil-COM/RtdManager.h>
#include <xlOil/RtdServer.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Caller.h>
#include <xlOil/Events.h>
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

  template<class A, class B>
  struct pair_hash {
    size_t operator()(std::pair<A, B> p) const noexcept
    {
      return xloil::boost_hash_combine(377, p.first, p.second);;
    }
  };
}

namespace xloil
{
  std::shared_ptr<IRtdServer> newRtdServer(
    const wchar_t* progId, const wchar_t* clsid)
  {
    return COM::newRtdServer(progId, clsid);
  }

  class AsyncTaskPublisher;

  struct CellTasks
  {
    //CellTasks(const msxll::XLREF12& caller_)
    //  : caller(caller_)
    //{}
    std::list<shared_ptr<AsyncTaskPublisher>> tasks;
    int arrayCount = 0; // see comment in 'getValue()'
    const wchar_t* arrayTopic = nullptr;
    msxll::XLREF12 caller;
    std::atomic_flag busy = ATOMIC_FLAG_INIT;
    bool isMaster(const msxll::XLREF12& ref) const
    {
      return ref.rwFirst == caller.rwFirst && ref.colFirst == caller.colFirst;
    }
    bool isSlave(const msxll::XLREF12& ref) const
    {
      return ref.rwFirst >= caller.rwFirst && ref.rwLast <= caller.rwLast
        && ref.colFirst >= caller.colFirst && ref.colLast <= caller.colLast;
    }

  };

  // TODO: could we just create a forwarding IRtdAsyncTask which intercepts 'cancel'
  class AsyncTaskPublisher : public RtdPublisher
  {
    std::atomic<bool> _disconnected = false;

  public:
    AsyncTaskPublisher(
      const wchar_t* topic,
      IRtdServer& mgr,
      const shared_ptr<IRtdTask>& task)
      : RtdPublisher(topic, mgr, task)
    {}

    bool disconnect(size_t numSubscribers) override
    {
      RtdPublisher::disconnect(numSubscribers);
      // TODO: check numSubscribers == 0
      stop();
      // Do nothing at this stage to avoid needing to lock CellTasks.
      // This flag will be checked by subsequent calls to rtdAsync.
      _disconnected = true;
      return true;
    }

    bool disconnected() const
    {
      return _disconnected;
    }

    bool taskMatches(const IRtdAsyncTask& that) const
    {
      return that == (const IRtdAsyncTask&)*task();
    }
  };

  class RtdAsyncManager
  {
  private:
    using CellAddress = std::pair<unsigned, unsigned>;
    shared_ptr<IRtdServer> _mgr;
    std::unordered_map<
      CellAddress,
      shared_ptr<CellTasks>,
      pair_hash<unsigned, unsigned>> _tasksPerCell;

    mutable std::shared_mutex _mutex;

    void start(decltype(CellTasks::tasks)& tasks, const shared_ptr<IRtdAsyncTask>& task) const
    {
      GUID guid;
      wchar_t guidStr[64];

      if (CoCreateGuid(&guid) != 0 || StringFromGUID2(guid, guidStr, _countof(guidStr)) == 0)
        XLO_THROW("Internal: RtdAsyncManager failed to create GUID");

      tasks.emplace_back(new AsyncTaskPublisher(guidStr, *_mgr, task));

      _mgr->start(tasks.back());
    }

    auto newCellTasks(
      const CellAddress& address, const msxll::XLREF12& ref)
    {
      //std::unique_lock writeLock(_mutex);
      // Emplace may not succeed if another thread has added the key whilst
      // we waited for the lock
      //const unsigned sheetId = address.first;
      auto [iTask, success] = _tasksPerCell.try_emplace(address, new CellTasks());
      auto* tasksInCell = &iTask->second;
      //if (success && (ref.colLast > ref.colFirst || ref.rwLast > ref.rwFirst))
      memcpy(&(*tasksInCell)->caller, &ref, sizeof(msxll::XLREF12));
      return tasksInCell;
    }

    void writeArray(
      const shared_ptr<CellTasks>& val,
      const unsigned sheetId, 
      const msxll::XLREF12& ref)
    {
      for (auto j = ref.colFirst; j <= ref.colLast; ++j)
        for (auto i = ref.rwFirst; i <= ref.rwLast; ++i)
        {
          unsigned num = i * XL_MAX_COLS + j;
          _tasksPerCell[make_pair(sheetId, num)] = val;
        }
    }

    RtdAsyncManager() : _mgr(newRtdServer())
    {
      // We're a singleton so guaranteed to still exist at autoclose
      Event::AutoClose() += [this]() { 
        clear(); 
        _mgr.reset(); 
      };
    }

  public:
    static RtdAsyncManager& instance()
    {
      static RtdAsyncManager* mgr = runExcelThread([]() 
        { 
          return new RtdAsyncManager(); 
        }).get();
      return *mgr;
    }

    shared_ptr<const ExcelObj> getValue(
      shared_ptr<IRtdAsyncTask> task)
    {
      // Protects agains a null-deref and allows starting up the RTD server
      // without running anything
      if (!task)
        return shared_ptr<const ExcelObj>();

      const auto caller = CallerLite();
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
      const auto sheetId = (unsigned)(intptr_t)caller.sheetId();
      const auto address = std::make_pair(sheetId, cellNumber);

      // The value we need to populate
      shared_ptr<CellTasks>* pTasksInCell;

      // Read-lock the dictionary of cell tasks and look for the address
      unique_lock writeLock(_mutex);
      const auto found = _tasksPerCell.find(address);
      if (found == _tasksPerCell.end())
      {
        // Not found, release read-lock and acquire write-lock
        //readLock.unlock();
        pTasksInCell = newCellTasks(address, *ref);
      }
      else
      {
        // Found: use the iterator *before* unlock as inserts may invalidate it
        pTasksInCell = &found->second;
        //readLock.unlock();
        
        auto* tasksInCell = pTasksInCell->get();

        if (!(*pTasksInCell)->isSlave(*ref))
        {
          pTasksInCell = newCellTasks(address, *ref);
        }
        else if (arraySize == 1 && tasksInCell->arrayCount > 0)
        {
          --tasksInCell->arrayCount;
          return _mgr->subscribe(tasksInCell->arrayTopic);
        }
        else
        {
          memcpy(&(tasksInCell->caller), ref, sizeof(msxll::XLREF12));
        }
      }

      if (arraySize > 1)
        writeArray(*pTasksInCell, sheetId, *ref);

      auto* tasksInCell = pTasksInCell->get();

      // (1) New master (no previous record)
      // (2) New master (former slave)
      //     (a) Shares top left
      //     (b) Doesn't
      // (3) Slave increment (arraySize = 1, arrayCount > 0)
   

      // If the caller is an array formula, when RTD is called in the subscribe()
      // method, it will return xlretUncalced, but will trigger the calling
      // function to be called again for each cell in the array. The caller
      // will remain as the top left-cell except for the first call which will 
      // be an array.
      // 
      // We want to start the task only once for the first call, with subsequent
      // calls ignored until the last one which must call subscribe (and hence 
      // RTD) for Excel to make the RTD connection.

      writeLock.unlock();
      // Now lock 'tasksInCell' in case there is more than one RTD function in the cell
      // This is unlikely, so we use a lightweight atomic_flag which implies a spin wait
      // (before C++20).
      scoped_atomic_flag lockCell(tasksInCell->busy);

      // Compare our task to all other running tasks in the cell to see if we
      // already have the answer

      shared_ptr<const ExcelObj> result;
      const wchar_t* arrayTopic;

      tasksInCell->tasks.remove_if([&](auto& t) {
        if (t->disconnected())
          return true;
        else if (!result && t->taskMatches(*task))
        {
          arrayTopic = t->topic();
          result = _mgr->peek(arrayTopic);
          // If we don't have a result or we got a result, but the task isn't 
          // done (e.g. it's a clock) then subscribe
          if (!result)
            result = _mgr->subscribe(arrayTopic);
          else if (!t->done())
            _mgr->subscribe(arrayTopic);
        }
        return false;
      });

      if (!result)
      {
        // Couldn't find a matching task so start a new one
        start(tasksInCell->tasks, task);
        //tasksInCell->arrayCount = arraySize;
        arrayTopic = tasksInCell->tasks.back()->topic();
        result = _mgr->subscribe(arrayTopic);
      }

      if (arraySize > 1)
      {
        tasksInCell->arrayCount = arraySize;
        tasksInCell->arrayTopic = arrayTopic;
      }
      else
      {
        tasksInCell->arrayCount = 0;
        tasksInCell->arrayTopic = nullptr;
      }

      return result;
    }

    void clear()
    {
      std::unique_lock lock(_mutex);
      _mgr->clear();
      _tasksPerCell.clear();
    }
  };

  shared_ptr<ExcelObj> rtdAsync(const shared_ptr<IRtdAsyncTask>& task)
  {
    // This cast is OK because if we are returning a non-null value we
    // will have cancelled the producer and nothing else will need the
    // ExcelObj
    return std::const_pointer_cast<ExcelObj>(
      RtdAsyncManager::instance().getValue(task));
  }

  void rtdAsyncServerClear()
  {
    RtdAsyncManager::instance().clear();
  }
}