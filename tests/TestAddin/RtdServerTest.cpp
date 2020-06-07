#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/Events.h>

using std::shared_ptr;

namespace xloil
{
  

  struct Counter
  {
    Counter(int iStep) : _iStep(iStep) {}

    int _iStep;

    std::future<void> operator()(IRtdNotify& notify)
    {
      return std::async([&notify, step = _iStep]()
      {
        int _count = 0;
        while (!notify.isCancelled())
        {
          notify.publish(ExcelObj(_count));
          std::this_thread::sleep_for(std::chrono::seconds(2));
          _count += step;
        }
      });
    }

    bool operator==(const Counter& that) const
    {
      return _iStep == that._iStep;
    }
  };


  XLO_FUNC_START(
    xloRtdCounter(const ExcelObj& step)
  )
  {
    auto iStep = step.toInt(1);
    auto value = rtdAsync(
      std::make_shared<RtdAsyncTask<Counter>>(iStep));
    return returnValue(value ? *value : CellError::NA);
  }
  XLO_FUNC_END(xloRtdCounter).macro();


  IRtdManager* getAnotherRtdServer()
  {
    static shared_ptr<IRtdManager> ptr = newRtdManager();
    // static auto deleter = Event::AutoClose() += [&]() { ptr.reset(); };
    return ptr.get();
  }

  XLO_FUNC_START(
    xloRtdSet(ExcelObj& tag, ExcelObj& val)
  )
  {
    auto topic = tag.toString();
    auto* srv = getAnotherRtdServer();
    if (!srv->peek(topic.c_str()).first)
      srv->start(std::make_shared<RtdTask>(
        [](IRtdNotify&) { return std::future<void>(); }), 
        topic.c_str(),
        true);
    srv->publish(topic.c_str(), ExcelObj(val));
    return returnValue(tag);
  }
  XLO_FUNC_END(xloRtdSet).macro();

  XLO_FUNC_START(
    xloRtdGet(ExcelObj& tag)
  )
  {
    auto value = getAnotherRtdServer()->subscribe(tag.toString().c_str());
    return returnReference(value
      ? *value
      : Const::Error(CellError::NA));
  }
  XLO_FUNC_END(xloRtdGet).macro();
}