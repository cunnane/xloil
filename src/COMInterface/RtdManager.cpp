#include "ComVariant.h"
#include "ClassFactory.h"

#include <xloil/ApiMessage.h>

#include <xloil/RtdServer.h>
#include <xloil/ExcelObj.h>
#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/Log.h>
#include <xloil/ExcelObjCache.h>
#include <xlOil/StringUtils.h>

#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <Objbase.h>

#include "ExcelTypeLib.h"

#include <unordered_map>
#include <unordered_set>
#include <memory>
#include <atomic>
#include <mutex>

using std::vector;
using std::shared_ptr;
using std::make_shared;
using std::scoped_lock;
using std::wstring;
using std::unique_ptr;
using std::future_status;
using std::unordered_set;
using std::unordered_map;
using std::pair;

namespace
{
  // ATL needs this for some reason
  class AtlModule : public CAtlDllModuleT<AtlModule>
  {} theAtlModule;
}

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
  void RtdPublisher::stop()
  {
    _task->cancel();
  }
  bool RtdPublisher::done() const
  {
    return _task->done();
  }
  const wchar_t* RtdPublisher::topic() const
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

  namespace COM
  {
    template<class TValue>
    class
      __declspec(novtable)
      RtdServerImpl :
        public CComObjectRootEx<CComSingleThreadModel>,
        public CComCoClass<RtdServerImpl<TValue>, &__uuidof(Excel::IRtdServer)>,
        public IDispatchImpl<
          Excel::IRtdServer,
          &__uuidof(Excel::IRtdServer),
          &LIBID_Excel>
    {
    public:

      ~RtdServerImpl()
      {
        try
        {
          clear();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RtdServer destruction: {0}", e.what());
        }
      }

      HRESULT _InternalQueryInterface(REFIID riid, void** ppv) throw()
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == __uuidof(Excel::IRtdServer))
        {
          *ppv = (IUnknown*)this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }

      HRESULT __stdcall raw_ServerStart(
        Excel::IRTDUpdateEvent* callback,
        long* result) override
      {
        if (!callback || !result)
          return E_POINTER;
        
        std::scoped_lock lock(_lockSubscribers);

        _updateCallback = callback;

        *result = 1;
        return S_OK;
      }

      HRESULT __stdcall raw_ConnectData(
        long topicId,
        SAFEARRAY** strings,
        VARIANT_BOOL* newValue,
        VARIANT* value) override
      {
        if (!strings || !newValue || !value)
          return E_POINTER;

        try
        {
          // We get the first topic string (and ignore the rest)
          long index[] = { 0, 0 };
          VARIANT topicAsVariant;
          SafeArrayGetElement(*strings, index, &topicAsVariant);

          const auto topic = topicAsVariant.bstrVal;

          std::scoped_lock lock(_lockSubscribers);

          // Find subscribers for this topic and link to the topic ID
          _activeTopicIds[topicId] = topic;
          auto& record = _records[topic];
          record.subscribers.insert(topicId);

          // Let the publisher know how many subscribers they now have
          if (record.publisher)
            record.publisher->connect(record.subscribers.size());

          XLO_DEBUG(L"RTD: connect '{}' to topicId '{}'", topic, topicId);
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd Disconnect: {0}", e.what());
        }
        return S_OK;
      }

      HRESULT __stdcall raw_RefreshData(
        long* topicCount,
        SAFEARRAY** data) override
      {
        decltype(_newValues) newValues;
        {
          // This lock is required for updates, so we avoid holding it and 
          // quickly splice out the list of new values. 
          // TODO: consider using lock-free list for updates
          scoped_lock lock(_lockNewValues);
          newValues.splice(newValues.begin(), _newValues);
        }

        scoped_lock lock(_lockSubscribers);
        unordered_set<long> readyTopicIds;
        for (auto&[topic, value] : newValues)
        {
          auto& record = _records[topic];
          record.value = value;
          readyTopicIds.insert(record.subscribers.begin(), record.subscribers.end());
        }

        const auto nReady = readyTopicIds.size();
        *topicCount = (long)nReady;

        if (nReady == 0)
          return S_OK; // There may be no subscribers for the producers

        SAFEARRAYBOUND bounds[] = { { 2u, 0 }, { (ULONG)nReady, 0 } };
        *data = SafeArrayCreate(VT_VARIANT, 2, bounds);

        void* element = nullptr;
        auto iRow = 0;
        for (auto topic : readyTopicIds)
        {
          long index[] = { 0, iRow };
          auto ret = SafeArrayPtrOfIndex(*data, index, &element);
          assert(S_OK == ret);
          *(VARIANT*)element = _variant_t(topic);

          index[0] = 1;
          ret = SafeArrayPtrOfIndex(*data, index, &element);
          assert(S_OK == ret);
          *(VARIANT*)element = _variant_t();

          ++iRow;
        }
        
        return S_OK;
      }

      HRESULT __stdcall raw_DisconnectData(long topicId) override
      {
        try
        {
          std::scoped_lock lock(_lockSubscribers);

          XLO_DEBUG("RTD: disconnect topicId {}", topicId);

          // Remove any done objects in the cancellation bucket
          _cancelledProducers.remove_if([](auto& x) { return x->done(); });

          const auto& topic = _activeTopicIds[topicId];
          if (topic.empty())
            XLO_THROW("Could not find topic for id {0}", topicId);

          auto& record = _records[topic];
          record.subscribers.erase(topicId);

          // If the disconnect() causes the publisher to cancel its task,
          // it will return true here. We may not be able to just delete it, 
          // we have to wait until any threads it created have exited
          if (record.publisher)
          {
            if (record.publisher->disconnect(record.subscribers.size()))
            {
              if (!record.publisher->done())
                cancelProducer(record.publisher);

              // Disconnect should only return true when num_subscribers = 0, 
              // so it's safe to erase the entire record
              _records.erase(topic);
            }
          }
          else if (record.subscribers.empty())
            _records.erase(topic);

          _activeTopicIds.erase(topicId);

        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd Disconnect: {0}", e.what());
        }

        return S_OK;
      }

      HRESULT __stdcall raw_Heartbeat(long* result) override
      {
        if (!result) return E_POINTER;
        *result = 1;
        return S_OK;
      }

      HRESULT __stdcall raw_ServerTerminate() override
      {
        try
        {
          if (!isServerRunning())
            return S_OK; // Already terminated, or never started

          // Terminate is called when there are no subscribers to the server
          // or the add-in is being closed. Clearing _updateCallback prevents 
          // the sending of further updates which could crash Excel.
          {
            scoped_lock lock(_lockNewValues);
            _updateCallback = nullptr;
            _newValues.clear();
          }

          clear();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd Disconnect: {0}", e.what());
        }
        return S_OK;
      }

    private:
      std::unordered_map<long, wstring> _activeTopicIds;

      struct TopicRecord
      {
        shared_ptr<IRtdPublisher> publisher;
        unordered_set<long> subscribers;
        shared_ptr<TValue> value;
      };
      
      unordered_map<wstring, TopicRecord> _records;

      std::list<pair<wstring, shared_ptr<TValue>>> _newValues;
      std::list<shared_ptr<IRtdPublisher>> _cancelledProducers;

      std::atomic<Excel::IRTDUpdateEvent*> _updateCallback;

      // We use a separate lock for the newValues to avoid blocking it 
      // too often: updates will come from other threads and just need to
      // write into newValues.
      mutable std::mutex _lockNewValues;
      mutable std::mutex _lockSubscribers;

      bool isServerRunning() const
      {
        return _updateCallback;
      }
      void cancelProducer(const shared_ptr<IRtdPublisher>& producer)
      {
        producer->stop();
        _cancelledProducers.push_back(producer);
      }

    public:
      void clear()
      {
        scoped_lock lock(_lockSubscribers);

        for (auto& record : _records)
        {
          try
          {
            if (record.second.publisher)
              record.second.publisher->stop();
          }
          catch (const std::exception& e)
          {
            XLO_INFO(L"Failed to stop producer: '{0}': {1}", 
              record.first, utf8ToUtf16(e.what()));
          }
        }

        _records.clear();
        _cancelledProducers.clear();
      }

      void update(const wchar_t* topic, const shared_ptr<TValue>& value)
      {
        if (!isServerRunning())
          return;

        scoped_lock lock(_lockNewValues);
        _newValues.push_back(make_pair(wstring(topic), value));

        // We only need to notify Excel about new data once. Excel will
        // only callback RefreshData approximately every 2 seconds
        if (_newValues.size() == 1)
        {
          excelApiCall([this]()
          {
            if (isServerRunning())
              (*_updateCallback).raw_UpdateNotify();
          }, QueueType::WINDOW, 100);
        }
      }

      void addProducer(const shared_ptr<IRtdPublisher>& job)
      {
        std::scoped_lock lock(_lockSubscribers);
        auto& record = _records[job->topic()];
        if (record.publisher)
          cancelProducer(record.publisher);
        record.publisher = job;
      }

      bool dropProducer(const wchar_t* topic)
      {
        std::scoped_lock lock(_lockSubscribers);
        auto i = _records.find(topic);
        if (i == _records.end())
          return false;

        // Signal the publisher to stop
        i->second.publisher->stop();

        // Destroy producer, the dtor of RtdPublisher waits for completion
        i->second.publisher.reset();

        // Publish empty value
        update(topic, shared_ptr<TValue>());
        return true;
      }
      bool value(const wchar_t* topic, shared_ptr<const TValue>& val) const
      {
        std::scoped_lock lock(_lockSubscribers);
        auto found = _records.find(topic);
        if (found == _records.end())
          return false;

        val = found->second.value;
        return true;
      }
    };

    class RtdServer : public IRtdServer
    {
      RegisterCom<RtdServerImpl<ExcelObj>> _registrar;
      RtdServerImpl<ExcelObj>* _server;
    public:
      RtdServer(const wchar_t* progId, const wchar_t* fixedClsid)
        : _registrar(progId, fixedClsid)
      {
        _server = &_registrar.server();
#ifdef _DEBUG
        //void* testObj;
        //res = CoCreateInstance(
        //  clsid, NULL,
        //  CLSCTX_INPROC_SERVER,
        //  __uuidof(Excel::IRtdServer),
        //  &testObj);
        //if (res != S_OK)
        //  XLO_ERROR(L"Failed to create com object '{0}'", _progId);
#endif
      }

      ~RtdServer()
      {
        clear();
      }

      void start(
        const shared_ptr<IRtdPublisher>& topic) override
      {
        _server->addProducer(topic);
      }

      shared_ptr<const ExcelObj> subscribe(const wchar_t * topic) override
      {
        callRtd(topic);
        shared_ptr<const ExcelObj> value;
        // If there is a producer, but no value yet, put N/A
        if (_server->value(topic, value) && !value)
          value = make_shared<ExcelObj>(CellError::NA);
        return value;
      }
      bool publish(const wchar_t * topic, ExcelObj&& value) override
      {
        _server->update(topic, make_shared<ExcelObj>(value));
        return true;
      }
      shared_ptr<const ExcelObj> 
        peek(const wchar_t* topic) override
      {
        shared_ptr<const ExcelObj> value;
        // If there is a producer, but no value yet, put N/A
        if (_server->value(topic, value) && !value)
          value = make_shared<ExcelObj>(CellError::NA);
        return value;
      }
      bool drop(const wchar_t* topic) override
      {
        return _server->dropProducer(topic);
      }
      const wchar_t* progId() const noexcept override
      {
        return _registrar.progid();
      }

      void clear() override
      {
        // This is likely be to called during teardown, so trap any errors
        try
        {
          // Although we can't force Excel to finalise the RTD server when we
          // want (as far as I know) we can deactive it and cut links to any 
          // external DLLs
          _server->clear();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RtdServer::clear: {0}", e.what());
        }
      }

    
      // We don't use the value from the Rtd call, but it indicates to
      // Excel that the current UDF should be treated as an RTD function.
      // It prompts Excel to connect to the RTD server to link a topicId
      // to this topic string
      void callRtd(const wchar_t* topic) const
      {
        callExcel(msxll::xlfRtd, _registrar.progid(), L"", topic);
      }
    };

    std::shared_ptr<IRtdServer> newRtdServer(
      const wchar_t* progId, const wchar_t* clsid)
    {
      if (!isMainThread())
        XLO_THROW("RtdServer must be created on main thread");
      return make_shared<COM::RtdServer>(progId, clsid);
    }
  }
}
