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

namespace xloil
{
  std::shared_ptr<IRtdManager> newRtdManager(
    const wchar_t* progId, const wchar_t* clsid)
  {
    return COM::newRtdManager(progId, clsid);
  }


  class RtdAsyncManager
  {
    using CellTasks = std::list<
      std::pair<std::shared_ptr<IRtdAsyncTask>, wstring>>;


    struct Cleanup : public IRtdAsyncTask
    {
      std::shared_ptr<IRtdAsyncTask> _task;
      CellTasks& _cellTasks;

      Cleanup(
        const shared_ptr<IRtdAsyncTask>& task,
        CellTasks& cellTasks)
        : _task(task)
        , _cellTasks(cellTasks)
      {}

      virtual ~Cleanup()
      {
        _cellTasks.remove_if(
          [task = _task](CellTasks::reference t)
        {
          return t.first == task;
        }
        );
      }

      std::future<void> operator()(IRtdNotify& n) override
      {
        return _task->operator()(n);
      }
      bool operator==(const IRtdAsyncTask& that) const override
      {
        return _task->operator==(that);
      }
    };

    std::unordered_map<wstring, CellTasks> _tasksPerCell;

    shared_ptr<IRtdManager> _mgr;

    void start(CellTasks& tasks, const shared_ptr<IRtdAsyncTask>& task)
    {
      GUID guid;
      CoCreateGuid(&guid);

      LPOLESTR guidStr;
      StringFromCLSID(guid, &guidStr);

      _mgr->start(make_shared<Cleanup>(task, tasks), guidStr, false);

      tasks.emplace_back(make_pair(task, guidStr));
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
      for (auto t = tasksInCell.begin(); t != tasksInCell.end(); ++t)
        if (*task == *t->first)
        {
          auto[value, isActive] = _mgr->peek(t->second.c_str());
          if (value && !isActive)
            return value;
          else
            return _mgr->subscribe(t->second.c_str());
        }

      // Couldn't find matching task so start it up
      start(tasksInCell, task);
      return _mgr->subscribe(tasksInCell.back().second.c_str());
    }
  };

  auto* getRtdAsyncManager()
  {
    // TODO: I guess we should create a mutux here, although calling
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
}