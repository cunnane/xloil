#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/Events.h>
//#include "Main.h"

using std::shared_ptr;

namespace xloil 
{
  namespace Test
  {
    struct Counter : public RtdAsyncTask
    {
      Counter(int iStep) : _iStep(iStep) {}

      int _iStep;

      std::future<void> operator()(RtdNotifier notify) override
      {
        return std::async([=, step = _iStep]()
        {
          int _count = 0;
          while (!notify.cancelled())
          {
            notify.publish(ExcelObj(_count));
            std::this_thread::sleep_for(std::chrono::seconds(2));
            _count += step;
          }
        });
      }

      bool operator==(const IRtdAsyncTask& that_) const override
      {
        const auto* that = dynamic_cast<const Counter*>(&that_);
        if (!that)
          return false;
        return _iStep == that->_iStep;
      }
    };

    XLO_FUNC_START(
      xloRtdCounter(const ExcelObj& step)
    )
    {
      auto iStep = step.toInt(1);
      auto value = rtdAsync(
        std::make_shared<Counter>(iStep));
      return returnValue(value ? *value : CellError::NA);
    }
    XLO_FUNC_END(xloRtdCounter);


    IRtdServer* getAnotherRtdServer()
    {
      static shared_ptr<IRtdServer> ptr = newRtdServer();
      //static auto deleter = Event_Shutdown() += [&]() { ptr.reset(); };
      return ptr.get();
    }

    XLO_FUNC_START(
      xloRtdSet(const ExcelObj& tag, const ExcelObj& val)
    )
    {
      auto topic = tag.toString();
      auto* srv = getAnotherRtdServer();
      if (!srv->peek(topic.c_str()))
        srv->start(topic.c_str(),
          [](RtdNotifier) { return std::future<void>(); });
      srv->publish(topic.c_str(), ExcelObj(val));
      return returnValue(tag);
    }
    XLO_FUNC_END(xloRtdSet);

    XLO_FUNC_START(
      xloRtdGet(const ExcelObj& tag)
    )
    {
      auto value = getAnotherRtdServer()->subscribe(tag.toString().c_str());
      return returnReference(value
        ? *value
        : Const::Error(CellError::NA));
    }
    XLO_FUNC_END(xloRtdGet);
  }
}