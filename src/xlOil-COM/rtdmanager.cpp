#include <xlOil/ExcelTypeLib.h>
#include "ComVariant.h"
#include "ClassFactory.h"

#include <xlOil/ExcelThread.h>

#include <xloil/RtdServer.h>
#include <xloil/ExcelObj.h>
#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/Log.h>
#include <xlOil/StringUtils.h>

#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <Objbase.h>

#include <unordered_map>
#include <unordered_set>
#include <memory>
#include <atomic>
#include <mutex>
#include <shared_mutex>

using std::vector;
using std::shared_ptr;
using std::make_shared;
using std::scoped_lock;
using std::unique_lock;
using std::shared_lock;
using std::wstring;
using std::unique_ptr;
using std::future_status;
using std::unordered_set;
using std::unordered_map;
using std::pair;
using std::list;

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

  namespace COM
  {
    template<class TValue>
    class
      __declspec(novtable)
      RtdServerImpl :
        public CComObjectRootEx<CComSingleThreadModel>,
        public CComCoClass<RtdServerImpl<TValue>>,
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
        
        unique_lock lock(_lockSubscribers);

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

          // We need these values after we release the lock
          shared_ptr<IRtdPublisher> publisher;
          size_t numSubscribers;
          {
            unique_lock lock(_lockSubscribers);

            // Find subscribers for this topic and link to the topic ID
            _activeTopicIds.emplace(topicId, topic);
            auto& record = _records[topic];
            record.subscribers.insert(topicId);
            publisher = record.publisher;
            numSubscribers = record.subscribers.size();
          }

          // Let the publisher know how many subscribers they now have.
          // We must not hold the lock when calling functions on the publisher
          // as they may try to call other functions on the RTD server. 
          if (publisher)
            publisher->connect(numSubscribers);

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

        unordered_set<long> readyTopicIds;
        {
          shared_lock lock(_lockSubscribers);
          for (auto&[topic, value] : newValues)
          {
            auto record = _records.find(topic);
            if (record == _records.end())
              continue;
            record->second.value = value;
            readyTopicIds.insert(record->second.subscribers.begin(), record->second.subscribers.end());
          }
        }

        const auto nReady = readyTopicIds.size();
        *topicCount = (long)nReady;

        // If no subscribers for the producers, we're done
        if (nReady == 0)
          return S_OK; 

        // 
        // All this code just creates a 2 x n safearray which has rows of:
        //     topicId | empty
        // With the topicId for each updated topic
        //
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
          // We will
          shared_ptr<IRtdPublisher> publisher;
          size_t numSubscribers;
          decltype(_cancelledPublishers) cancelledPublishers;

          // We must *not* hold the lock when calling methods of the publisher
          // as they may try to call other functions on the RTD server. So we
          // first handle the topic lookup and removing subscribers before
          // releasing the lock and notifying the publisher.
          {
            unique_lock lock(_lockSubscribers);

            XLO_DEBUG("RTD: disconnect topicId {}", topicId);

            std::swap(_cancelledPublishers, cancelledPublishers);

            const auto topic = _activeTopicIds.find(topicId);
            if (topic == _activeTopicIds.end())
              XLO_THROW("Could not find topic for id {0}", topicId);

            auto& record = _records[topic->second];
            record.subscribers.erase(topicId);

            numSubscribers = record.subscribers.size();
            publisher = record.publisher;

            if (!publisher && numSubscribers == 0)
              _records.erase(topic->second);

            _activeTopicIds.erase(topic);
          }

          // Remove any done objects in the cancellation bucket
          cancelledPublishers.remove_if([](auto& x) { return x->done(); });

          if (!publisher)
            return S_OK;

          // If the disconnect() causes the publisher to cancel its task,
          // it will return true here. We may not be able to just delete it, 
          // we have to wait until any threads it created have exited
          if (publisher->disconnect(numSubscribers))
          {
            const auto topic = publisher->topic();
            const auto done = publisher->done();

            if (!done)
              publisher->stop();
            {
              unique_lock lock(_lockSubscribers);

              if (!done)
                _cancelledPublishers.emplace_back(publisher);

              // Disconnect should only return true when num_subscribers = 0, 
              // so it's safe to erase the entire record
              _records.erase(topic);
            }
          }
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

      list<pair<wstring, shared_ptr<TValue>>> _newValues;
      std::future<void> _updateNotifyTask;
      list<shared_ptr<IRtdPublisher>> _cancelledPublishers;

      std::atomic<Excel::IRTDUpdateEvent*> _updateCallback;

      // We use a separate lock for the newValues to avoid blocking too 
      // often: value updates are likely to come from other threads and 
      // simply need to write into newValues without accessing pub/sub info.
      // We use _lockSubscribers for all other synchronisation
      mutable std::mutex _lockNewValues;
      mutable std::shared_mutex _lockSubscribers;

      bool isServerRunning() const
      {
        return _updateCallback;
      }

    public:
      void clear()
      {
        // We must not hold the lock when calling functions on the publisher
        // as they may try to call other functions on the RTD server. 

        list<shared_ptr<IRtdPublisher>> publishers;
        {
          unique_lock lock(_lockSubscribers);

          for (auto& record : _records)
            if (record.second.publisher)
              publishers.emplace_back(std::move(record.second.publisher));

          _records.clear();
          _cancelledPublishers.clear();
        }
        
        for (auto& pub : publishers)
        {
          try
          {
            pub->stop();
          }
          catch (const std::exception& e)
          {
            XLO_INFO(L"Failed to stop producer: '{0}': {1}",
              pub->topic(), utf8ToUtf16(e.what()));
          }
        }
      }

      void update(const wchar_t* topic, const shared_ptr<TValue>& value)
      {
        if (!isServerRunning())
          return;

        scoped_lock lock(_lockNewValues);

        _newValues.push_back(make_pair(wstring(topic), value));

        // We only need to notify Excel about new data once, so check if
        // our notify task has successfully completed. Excel will
        // only callback RefreshData every 2 seconds (unless someone fiddled 
        // with the throttle interval)
        if (!_updateNotifyTask.valid() || _updateNotifyTask._Is_ready())
        {
          // Must Enqueue the update notify call as if it is called from another RTD
          // entry point, Excel will ignore it. 1 sec between retries.
          _updateNotifyTask = runExcelThread([this]()
          {
            if (isServerRunning())
              (*_updateCallback).raw_UpdateNotify(); // Does this really need the COM API?
          }, ExcelRunQueue::COM_API | ExcelRunQueue::ENQUEUE, 0, 1000); 
        }
      }

      void addProducer(shared_ptr<IRtdPublisher> job)
      {
        {
          unique_lock lock(_lockSubscribers);
          auto& record = _records[job->topic()];
          std::swap(record.publisher, job);
          if (job)
            _cancelledPublishers.push_back(job);
        }
        if (job)
          job->stop();
      }

      bool dropProducer(const wchar_t* topic)
      {
        // We must not hold the lock when calling functions on the publisher
        // as they may try to call other functions on the RTD server. 
        shared_ptr<IRtdPublisher> publisher;
        {
          unique_lock lock(_lockSubscribers);
          auto i = _records.find(topic);
          if (i == _records.end())
            return false;
          std::swap(publisher, i->second.publisher);
        }

        // Signal the publisher to stop
        publisher->stop();

        // Destroy producer, the dtor of RtdPublisher waits for completion
        publisher.reset();

        // Publish empty value
        update(topic, shared_ptr<TValue>());
        return true;
      }
      bool value(const wchar_t* topic, shared_ptr<const TValue>& val) const
      {
        shared_lock lock(_lockSubscribers);
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
      RtdServerImpl<ExcelObj>& server() const { return *_registrar.server(); }

    public:
      RtdServer(const wchar_t* progId, const wchar_t* fixedClsid)
        : _registrar(
          [](const wchar_t*, const GUID&) { return new CComObject<RtdServerImpl<ExcelObj>>(); },
          progId, 
          fixedClsid)
      {
        // TODO: why doesn't this work?
        //_registrar.cleanRegistry();
      }

      ~RtdServer()
      {
        clear();
      }

      void start(
        const shared_ptr<IRtdPublisher>& topic) override
      {
        server().addProducer(topic);
      }

      shared_ptr<const ExcelObj> subscribe(const wchar_t * topic) override
      {
        shared_ptr<const ExcelObj> value;
        // If there is a producer, but no value yet, put N/A
        callRtd(topic);
        if (server().value(topic, value) && !value)
          value = make_shared<ExcelObj>(CellError::NA);
        return value;
      }
      bool publish(const wchar_t * topic, ExcelObj&& value) override
      {
        server().update(topic, make_shared<ExcelObj>(value));
        return true;
      }
      shared_ptr<const ExcelObj> 
        peek(const wchar_t* topic) override
      {
        shared_ptr<const ExcelObj> value;
        // If there is a producer, but no value yet, put N/A
        if (server().value(topic, value) && !value)
          value = make_shared<ExcelObj>(CellError::NA);
        return value;
      }
      bool drop(const wchar_t* topic) override
      {
        return server().dropProducer(topic);
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
          server().clear();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RtdServer::clear: {0}", e.what());
        }
      }
    
      /// <summary>
      /// We don't use the value from the Rtd call, but it indicates to
      /// Excel that the current UDF should be treated as an RTD function.
      /// It prompts Excel to connect to the RTD server to link a topicId
      /// to this topic string. Returns false if the excel call fails 
      /// (usually this will be due to xlretUncalced which occurs when the
      /// calling cell is an array formula).
      /// </summary>
      bool callRtd(const wchar_t* topic) const
      {
        auto[val, retCode] = 
          tryCallExcel(msxll::xlfRtd, _registrar.progid(), L"", topic);
        return retCode == 0;
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
