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
using std::scoped_lock;

namespace
{
  /// <summary>
  /// Like a std::scoped_lock but uses a std::atomic_flag rather than a mutex.
  /// Note it busy-waits for the lock!
  /// </summary>
  struct scoped_atomic_flag
  {
    std::atomic_flag& _flag;

    scoped_atomic_flag(std::atomic_flag& flag)
      : _flag(flag)
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
      _flag.clear(std::memory_order_release);
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
  using CellTasks = std::list<shared_ptr<AsyncTaskPublisher>>;

  // TODO: could we just create a forwarding IRtdAsyncTask which intercepts 'cancel'
  class AsyncTaskPublisher : public RtdPublisher
  {
    CellTasks& _cellTasks;

  public:
    AsyncTaskPublisher(
      const wchar_t* topic,
      IRtdServer& mgr,
      const shared_ptr<IRtdTask>& task,
      CellTasks& cellTasks)
      : RtdPublisher(topic, mgr, task)
      , _cellTasks(cellTasks)
    {}

    bool disconnect(size_t numSubscribers) override
    {
      RtdPublisher::disconnect(numSubscribers);
      // TODO: check numSubscribers == 0
      stop();
      _cellTasks.remove_if(
        [target = this](CellTasks::reference t)
        {
          return t.get() == target;
        }
      );
      return true;
    }
  };

  class RtdAsyncManager
  {
  private:
    struct CellTaskHolder
    {
      CellTasks tasks;
      int arrayCount = 0; // see comment in 'getValue()'
      std::atomic_flag busy = ATOMIC_FLAG_INIT;
    };

    void start(CellTasks& tasks, const shared_ptr<IRtdAsyncTask>& task)
    {
      GUID guid;
      wchar_t guidStr[64];

      if (CoCreateGuid(&guid) != 0 || StringFromGUID2(guid, guidStr, _countof(guidStr)) == 0)
        XLO_THROW("Internal: RtdAsyncManager failed to create GUID");

      tasks.emplace_back(new AsyncTaskPublisher(guidStr, *_mgr, task, tasks));

      _mgr->start(tasks.back());
    }

  private:
    shared_ptr<IRtdServer> _mgr;
    std::unordered_map<std::pair<unsigned, unsigned>, CellTaskHolder, pair_hash<unsigned, unsigned>> _tasksPerCell;
    std::shared_mutex _mutex;

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
      static RtdAsyncManager* mgr = runExcelThread([]() { return new RtdAsyncManager(); }).get();
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
      const auto arraySize = (ref->colLast - ref->colFirst + 1) 
        * (ref->rwLast - ref->rwFirst + 1);
      // This is the cell number of the top-left cell for array callers
      const unsigned cellNumber = ref->rwFirst * XL_MAX_COLS + ref->colFirst;
      const auto address = std::make_pair((unsigned)caller.sheetId(), cellNumber);

      // The value we need to populate
      CellTaskHolder* tasksInCell;

      // Read-lock the dictionary of cell tasks and look for the address
      std::shared_lock readLock(_mutex);
      const auto iTasks = _tasksPerCell.find(address);
      if (iTasks == _tasksPerCell.end())
      {
        // Not found, release read-lock and acquire write-lock
        readLock.unlock();
        
        std::unique_lock writeLock(_mutex);
        // Emplace may not succeed if another thread has added the key whilst
        // we waited for the lock
        tasksInCell = &_tasksPerCell.try_emplace(address).first->second;
      }
      else
      {
        // Found: use the iterator *before* unlock as inserts may invalidate it
        tasksInCell = &iTasks->second;
        readLock.unlock();
      }
      
      // Now lock 'tasksInCell' in case there is more than one RTD function in the cell
      // This is unlikely, so we use a lightweight atomic_flag which implies a spin wait
      // (before C++20).
      scoped_atomic_flag lockCell(tasksInCell->busy);

      // If the caller is an array formula, when RTD is called in the subscribe()
      // method, it will return xlretUncalced, but will trigger the calling
      // function to be called again for each cell in the array. The caller
      // will remain as the top left-cell except for the first call which will 
      // be an array.
      // 
      // We want to start the task only once for the first call, with subsequent
      // calls ignored until the last one which must call subscribe (and hence 
      // RTD) for Excel to make the RTD connection.
      if (tasksInCell->arrayCount > 1)
      {
        --tasksInCell->arrayCount;
        return shared_ptr<const ExcelObj>();
      }

      // Compare our task to all other running tasks in the cell to see if we
      // already have the answer
      for (auto& t: tasksInCell->tasks)
        if (*task == (const IRtdAsyncTask&)*t->task())
        {
          auto value = _mgr->peek(t->topic());
          return value && t->done()
            ? value 
            : _mgr->subscribe(t->topic());
        }

      // Couldn't find a matching task so start a new one
      start(tasksInCell->tasks, task);
      tasksInCell->arrayCount = arraySize;
      auto result = _mgr->subscribe(tasksInCell->tasks.back()->topic());
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