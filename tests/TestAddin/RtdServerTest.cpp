#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/Events.h>

using std::shared_ptr;

namespace xloil
{
  XLO_FUNC_START(
    xloRtdCounter()
  )
  {
    return returnReference(rtdConnect().run(
        [](IRtdNotify& notify)
        {
          return std::async([&]() 
          {
            int _count = 0;
            while (!notify.isCancelled())
            {
              notify.publish(ExcelObj(_count));
              std::this_thread::sleep_for(std::chrono::seconds(2));
              ++_count;
            }
          });
        })
    );
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
    if (!srv->peek(topic.c_str()))
      srv->start(
        [](IRtdNotify&) { return std::future<void>(); }, 
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
    auto conn = rtdConnect(getAnotherRtdServer(), tag.toString().c_str());
    return returnReference(conn.hasValue()
      ? conn.value()
      : Const::Error(CellError::NA));
  }
  XLO_FUNC_END(xloRtdGet).macro();
}