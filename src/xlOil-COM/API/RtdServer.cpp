#include <xlOil-COM/RtdManager.h>
#include <xlOil-COM/RtdAsyncManager.h>
#include <xlOil/RtdServer.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Caller.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelThread.h>
#include <xloil/StringUtils.h>
#include <combaseapi.h>

using std::wstring;
using std::shared_ptr;
using std::unique_ptr;
using std::make_shared;
using std::make_pair;


namespace xloil
{
  RtdPublisher::RtdPublisher(
    const wchar_t* topic,
    IRtdServer& mgr,
    const shared_ptr<IRtdTask>& task)
    : _mgr(mgr)
    , _task(task)
    , _topic(topic)
  {}

  RtdPublisher::~RtdPublisher()
  {
    try
    {
      // Send cancellation and wait for graceful shutdown
      stop();
      _task->wait();
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Rtd Disconnect: {0}", e.what());
    }
  }

  void RtdPublisher::connect(size_t numSubscribers)
  {
    if (numSubscribers == 1)
    {
      _task->start(*this);
    }
  }
  bool RtdPublisher::disconnect(size_t numSubscribers)
  {
    if (numSubscribers == 0)
    {
      stop();
      return true;
    }
    return false;
  }
  void RtdPublisher::stop() noexcept
  {
    _task->cancel();
  }
  bool RtdPublisher::done() const noexcept
  {
    return _task->done();
  }
  const wchar_t* RtdPublisher::topic() const noexcept
  {
    return _topic.c_str();
  }
  bool RtdPublisher::publish(ExcelObj&& value) noexcept
  {
    try
    {
      _mgr.publish(_topic.c_str(), std::forward<ExcelObj>(value));
      return true;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR(L"RTD error publishing {}: {}", _topic, utf8ToUtf16(e.what()));
    }
    return false;
  }

  std::shared_ptr<IRtdServer> newRtdServer(
    const wchar_t* progId, const wchar_t* clsid)
  {
    return COM::newRtdServer(progId, clsid);
  }

  shared_ptr<ExcelObj> rtdAsync(const shared_ptr<IRtdAsyncTask>& task)
  {
    // This cast is OK because if we are returning a non-null value we
    // will have cancelled the producer and nothing else will need the
    // ExcelObj
    return std::const_pointer_cast<ExcelObj>(
      COM::RtdAsyncManager::getValue(task));
  }

  void rtdAsyncServerClear()
  {
    COM::RtdAsyncManager::clear();
  }
}