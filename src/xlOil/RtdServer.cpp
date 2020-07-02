#include <COMinterface/RtdManager.h>
#include <xloil/RtdServer.h>
#include <xlOilHelpers/WindowsSlim.h>
#include <xloil/Caller.h>
#include <xlOil/Events.h>
#include <combaseapi.h>

using std::wstring;
using std::shared_ptr;
using std::unique_ptr;
using std::make_shared;
using std::scoped_lock;

namespace xloil
{
  std::shared_ptr<IRtdManager> newRtdManager(
    const wchar_t* progId, const wchar_t* clsid)
  {
    return COM::newRtdManager(progId, clsid);
  }

  class AsyncTaskTopic;
  using CellTasks = std::list<std::shared_ptr<AsyncTaskTopic>>;

  // TODO: could we just create a forwarding IRtdAsyncTask which intercepts 'cancel'
  class AsyncTaskTopic : public RtdTopic
  {
    CellTasks& _cellTasks;

  public:
    AsyncTaskTopic(
      const wchar_t* topic,
      IRtdManager& mgr,
      const shared_ptr<IRtdProducer>& task,
      CellTasks& cellTasks)
      : RtdTopic(topic, mgr, task)
      , _cellTasks(cellTasks)
    {
    }

    bool disconnect(size_t numSubscribers) override
    {
      RtdTopic::disconnect(numSubscribers);
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
    shared_ptr<IRtdManager> _mgr;
    std::unordered_map<wstring, CellTasks> _tasksPerCell;

    void start(CellTasks& tasks, const shared_ptr<IRtdAsyncTask>& task)
    {
      GUID guid;
      CoCreateGuid(&guid);

      LPOLESTR guidStr;
      StringFromCLSID(guid, &guidStr);

      auto topic = make_shared<AsyncTaskTopic>(guidStr, *_mgr, task, tasks);
      _mgr->start(topic);

      tasks.emplace_back(topic);
      CoTaskMemFree(guidStr);
    }

  public:
    RtdAsyncManager() : _mgr(newRtdManager())
    {}

    shared_ptr<const ExcelObj> getValue(
      shared_ptr<IRtdAsyncTask> task)
    {
      const auto address = CallerInfo().writeAddress(false);
      auto& tasksInCell = _tasksPerCell[address];
      for (auto& t: tasksInCell)
        if (*task == (const IRtdAsyncTask&)*t->task())
        {
          auto value = _mgr->peek(t->topic());
          return value && t->done()
            ? value 
            : _mgr->subscribe(t->topic());
        }

      // Couldn't find matching task so start it up
      start(tasksInCell, task);
      return _mgr->subscribe(tasksInCell.back()->topic());
    }

    void clear()
    {
      for (auto& cell : _tasksPerCell)
      {
        if (!cell.second.empty())
        {
          // Make a copy of the list as the dtor of Cleanup edits it
          auto copy(cell.second);
          for (auto& j : copy)
            _mgr->drop(j->topic());
        }
      }
       
      _tasksPerCell.clear();
    }
  };

  auto* getRtdAsyncManager()
  {
    // TODO: I guess we should create a mutex here, although calling
    // RTD functions from a multithreaded function is not likely to 
    // end well. Can we check for that in a non-expensive way?
    static auto ptr = make_shared<RtdAsyncManager>();
    static auto deleter = Event::AutoClose() += [&]() { ptr.reset(); };
    return ptr.get();
  }

  shared_ptr<ExcelObj> rtdAsync(const shared_ptr<IRtdAsyncTask>& task)
  {
    // This cast is OK because if we are returning a non-null value we
    // will have cancelled the producer and nothing else will need the
    // ExcelObj
    return std::const_pointer_cast<ExcelObj>(
      getRtdAsyncManager()->getValue(task));
  }

  void rtdAsyncManagerClear()
  {
    getRtdAsyncManager()->clear();
  }
}