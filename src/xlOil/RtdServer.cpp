#include <COMinterface/RtdManager.h>
#include <xloil/RtdServer.h>
#include <xlOil/WindowsSlim.h>
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
    shared_ptr<IRtdServer> _mgr;
    std::unordered_map<wstring, CellTasks> _tasksPerCell;

    void start(CellTasks& tasks, const shared_ptr<IRtdAsyncTask>& task)
    {
      GUID guid;
      CoCreateGuid(&guid);

      LPOLESTR guidStr;
      StringFromCLSID(guid, &guidStr);

      auto topic = make_shared<AsyncTaskPublisher>(guidStr, *_mgr, task, tasks);
      CoTaskMemFree(guidStr);

      _mgr->start(topic);

      tasks.emplace_back(topic);
    }

  public:
    RtdAsyncManager() : _mgr(newRtdServer())
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
      _mgr->clear();
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

  void rtdAsyncServerClear()
  {
    getRtdAsyncManager()->clear();
  }
}