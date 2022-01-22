#include <xlOil-COM/RtdManager.h>
#include <xlOil/RtdServer.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Caller.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelThread.h>
#include <combaseapi.h>

using std::wstring;
using std::shared_ptr;
using std::unique_ptr;
using std::make_shared;
using std::scoped_lock;

namespace
{

  /// <summary>
  /// Like a std::scoped_lock but uses a std::atomic_flag rather than a mutex
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
    std::unordered_map<wstring, CellTaskHolder> _tasksPerCell;
    std::mutex _lock;

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
      // If the caller is an array formula, when RTD is called in the subscribe()
      // method, it will return xlretUncalced, but will trigger the calling
      // function to be called again for each cell in the array. The caller
      // will remain as the top left-cell except for the first call which will 
      // be an array.
      // 
      // We want to start the task only once for the first call, with subsequent
      // calls ignored until the last one which must call subscribe (and hence 
      // RTD) for Excel to make the RTD connection.
      //
      if (!task)
        return shared_ptr<const ExcelObj>();

      const auto caller = CallerInfo();
      const auto ref = caller.sheetRef();
      const auto arraySize = (ref->colLast - ref->colFirst + 1) 
        * (ref->rwLast - ref->rwFirst + 1);
      auto address = caller.writeAddress(CallerInfo::RC);

      // Turn array address into top left cell
      if (arraySize > 1)
        address.erase(address.rfind(':'));
  
      // TODO: shared_mutex here
      _lock.lock();
      auto& tasksInCell = _tasksPerCell[address];
      scoped_atomic_flag lockCell(tasksInCell.busy);
      _lock.unlock();

      if (tasksInCell.arrayCount > 1)
      {
        --tasksInCell.arrayCount;
        return shared_ptr<const ExcelObj>();
      }

      for (auto& t: tasksInCell.tasks)
        if (*task == (const IRtdAsyncTask&)*t->task())
        {
          auto value = _mgr->peek(t->topic());
          return value && t->done()
            ? value 
            : _mgr->subscribe(t->topic());
        }

      // Couldn't find matching task so start it up
      start(tasksInCell.tasks, task);
      tasksInCell.arrayCount = arraySize;
      auto result = _mgr->subscribe(tasksInCell.tasks.back()->topic());
      return result;
    }

    void clear()
    {
      std::scoped_lock lock(_lock);
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