#include "ComVariant.h"
#include <xloil/ThreadControl.h>

#include <xloil/RtdServer.h>
#include <xloil/ExcelObj.h>
#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/Log.h>
#include <xloil/ExcelObjCache.h>
#include <xlOilHelpers/StringUtils.h>

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

namespace
{
  // ATL needs this for some reason
  class AtlModule : public CAtlDllModuleT<AtlModule>
  {} theAtlModule;
}

namespace xloil 
{
  

  RtdTopic::RtdTopic(
    const wchar_t* topic,
    IRtdManager& mgr, 
    const shared_ptr<IRtdProducer>& task)
    : _mgr(mgr)
    , _task(task)
    , _topic(topic)
  {}

  RtdTopic::~RtdTopic()
  {
    // Send cancellation and wait for graceful shutdown
    stop();
    _task->wait();
  }

  void RtdTopic::connect(size_t numSubscribers)
  {
    if (numSubscribers == 1)
    {
      _task->start(*this);
    }
  }
  bool RtdTopic::disconnect(size_t numSubscribers)
  {
    if (numSubscribers == 0)
    {
      stop();
      return true;
    }
    return false;
  }
  void RtdTopic::stop()
  {
    _task->cancel();
  }
  bool RtdTopic::done() const
  {
    return _task->done();
  }
  const wchar_t* RtdTopic::topic() const
  {
    return _topic.c_str();
  }
  bool RtdTopic::publish(ExcelObj&& value) noexcept
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
    class
      __declspec(novtable)
      RtdServerImpl :
        public CComObjectRootEx<CComSingleThreadModel>,
        public CComCoClass<RtdServerImpl, &__uuidof(Excel::IRtdServer)>,
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
          for (auto& prod : _theProducers)
            prod.second->stop();
          _theProducers.clear();
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
          // We only look at the first topic string
          long index[] = { 0, 0 };
          VARIANT topicVariant;
          SafeArrayGetElement(*strings, index, &topicVariant);
          const auto topic = topicVariant.bstrVal;

          std::scoped_lock lock(_lockSubscribers);

          _activeTopicIds[topicId] = topic;
          auto& subscribers = _subscribers[topic];
          subscribers.insert(topicId);
          size_t nSubscribers = subscribers.size();

          auto producer = findProducer(topic);
          if (producer)
            producer->connect(nSubscribers);

          XLO_TRACE(L"RTD: connect '{}' to topicId '{}'", topic, topicId);
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
        std::scoped_lock lock(_lockSubscribers);

        unordered_set<long> readyTopicIds;
        {
          std::scoped_lock lock(_lockNewValues);
          for (auto&[topic, value] : _newValues)
          {
            _values[topic] = value;
            const auto& subscribers = _subscribers[topic];
            readyTopicIds.insert(subscribers.begin(), subscribers.end());
          }

          _newValues.clear();
        }

        const auto nReady = readyTopicIds.size();
        *topicCount = (long)nReady;

        if (nReady == 0)
          return S_OK; // There may be no subscribers for the producers

        SAFEARRAYBOUND bounds[] = { { 2u, 0 }, { (ULONG)nReady, 0 } };
        *data = SafeArrayCreate(VT_VARIANT, 2, bounds);

        void* element;
        auto iRow = 0;
        for (auto topic : readyTopicIds)
        {
          long index[] = { 0, iRow };
          assert(S_OK == SafeArrayPtrOfIndex(*data, index, &element));
          *(VARIANT*)element = _variant_t(topic);

          index[0] = 1;
          assert(S_OK == SafeArrayPtrOfIndex(*data, index, &element));
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

          XLO_TRACE("RTD: disconnect topicId {}", topicId);

          // Remove any done objects in the cancellation bucket
          _cancelledProducers.remove_if([](auto& x) { return x->done(); });

          const auto& topic = _activeTopicIds[topicId];
          if (topic.empty())
            XLO_THROW("Could not find topic for id {0}", topicId);

          auto& subscribers = _subscribers[topic];
          subscribers.erase(topicId);

          // If the disconnect() causes the producer to cancel its task,
          // it will return true here. We may not be able to just delete it, 
          // we have to wait until any threads it created exit
          auto producer = findProducer(topic.c_str());
          if (producer && producer->disconnect(subscribers.size()))
          {
            if (!producer->done())
              cancelProducer(producer);
            _theProducers.erase(topic);
          }

          if (subscribers.empty())
            _subscribers.erase(topic);
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
        std::scoped_lock lock(_lockSubscribers);
        try
        {
          for (auto& prod : _theProducers)
            prod.second->stop();
          _activeTopicIds.clear();
          _updateCallback = nullptr;
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd Disconnect: {0}", e.what());
        }
        return S_OK;
      }

    private:
      std::unordered_map<long, wstring> _activeTopicIds;

      // TODO: combine these maps?
      std::unordered_map<wstring, shared_ptr<IRtdTopic>> _theProducers;
      std::unordered_map<wstring, unordered_set<long>> _subscribers;
      std::unordered_map<wstring, shared_ptr<ExcelObj>> _values;
      std::unordered_map<wstring, shared_ptr<ExcelObj>> _newValues;

      std::list<shared_ptr<IRtdTopic>> _cancelledProducers;

      Excel::IRTDUpdateEvent* _updateCallback;
      // We use a separate lock for the newValues to avoid blocking it 
      // too often: updates will come from other threads and just need to
      // write into newValues.
      std::mutex _lockNewValues;
      std::mutex _lockSubscribers;

      void callUpdateNotify()
      {
        std::scoped_lock lock(_lockSubscribers);
        if (_updateCallback)
          _updateCallback->UpdateNotify();
      }

      shared_ptr<IRtdTopic> findProducer(const wchar_t* topic)
      {
        auto iJob = _theProducers.find(topic);
        if (iJob == _theProducers.end())
          return nullptr;
        return iJob->second;
      }

      void cancelProducer(const shared_ptr<IRtdTopic>& producer)
      {
        producer->stop();
        _cancelledProducers.push_back(producer);
      }

    public:
      void update(const wchar_t* topic, const shared_ptr<ExcelObj>& value)
      {
        // TODO: could write this non-blocking by using a singly linked list of values
        std::scoped_lock lock(_lockNewValues);

        _newValues[topic] = value;

        // We only need to notify Excel about new data once. Excel will
        // only callback RefreshData approximately every 2 seconds
        if (_newValues.size() == 1 && _updateCallback)
        {
          queueWindowMessage([this]() { this->callUpdateNotify(); });
        }
      }

      void addProducer(const shared_ptr<IRtdTopic>& job)
      {
        std::scoped_lock lock(_lockSubscribers);
        // TODO: can we make map key wchar_t*??
        auto iProd = _theProducers.find(job->topic());
        if (iProd != _theProducers.end())
        {
          cancelProducer(iProd->second);
          iProd->second = job;
        }
        else
          _theProducers.emplace(job->topic(), job);
      }

      bool dropProducer(const wchar_t* topic)
      {
        
        std::scoped_lock lock(_lockSubscribers);
        auto iProd = _theProducers.find(topic);
        if (iProd == _theProducers.end())
          return false;

        // Signal the producer to stop
        iProd->second->stop();
        // Destroy producer, the dtor of RtdTopic waits for completion
        _theProducers.erase(iProd);

        _values.erase(topic); // TODO: does this throw if not found?
        
        return true;
      }

      auto value(const wchar_t* topic)
      {
        std::scoped_lock lock(_lockSubscribers);
        return _values[topic];
      }
    };

    class FactoryRtdServer : public IClassFactory
    {
    public:
      RtdServerImpl* _instance;

      FactoryRtdServer(RtdServerImpl* p)
        : _instance(p)
      {}

      HRESULT STDMETHODCALLTYPE CreateInstance(
        IUnknown *pUnkOuter,
        REFIID riid,
        void **ppvObject) override
      {
        if (pUnkOuter)
          return CLASS_E_NOAGGREGATION;
        auto ret = _instance->QueryInterface(riid, ppvObject);
        return ret;
      }

      HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override
      {
        return E_NOTIMPL;
      }

      HRESULT QueryInterface(REFIID riid, void** ppv)
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == __uuidof(IClassFactory))
        {
          *ppv = (IUnknown*)this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }

      STDMETHOD_(ULONG, AddRef)() override
      {
        InterlockedIncrement(&_cRef);
        return _cRef;
      }

      STDMETHOD_(ULONG, Release)() override
      {
        InterlockedDecrement(&_cRef);
        if (_cRef == 0)
        {
          delete this;
          return 0;
        }
        return _cRef;
      }
    private:
      LONG _cRef;
    };

    HRESULT regWriteString(
      HKEY hive,
      const wchar_t* path,
      const wchar_t* name,
      const wchar_t* value)
    {
      HRESULT res;
      HKEY key;

      if (0 > (res = RegCreateKeyExW(
        hive,
        path,
        0, NULL,
        REG_OPTION_VOLATILE, // key not saved on system shutdown
        KEY_ALL_ACCESS,      // no access restrictions
        NULL,                // no security restrictions
        &key, NULL)))
        return res;

      res = RegSetValueEx(
        key,
        name,
        0,
        REG_SZ, // String type
        (BYTE*)value,
        (DWORD)(wcslen(value) + 1) * sizeof(wchar_t));

      RegCloseKey(key);

      return res;
    }

    class RtdManager : public IRtdManager
    {
      CComPtr<RtdServerImpl> _server;
      CComPtr<FactoryRtdServer> _factory;
      DWORD _comRegistrationCookie;
      wstring _progId;
      std::list<wstring> _regKeysAdded;

    public:
      RtdManager(const wchar_t* progId, const wchar_t* fixedClsid)
      {
        _server = new CComObject<RtdServerImpl>();
        _factory = new FactoryRtdServer(_server.p);

        if (progId && !fixedClsid)
          XLO_THROW("If you specify an RTD ProgId you must also specify a "
            "CLSID or different Excel instances may clash");

        GUID clsid;
        HRESULT hCreateGuid = fixedClsid
          ? CLSIDFromString(fixedClsid, &clsid)
          : CoCreateGuid(&clsid);

        LPOLESTR clsidStr;
        // This generates the string '{W-X-Y-Z}'
        StringFromCLSID(clsid, &clsidStr);

        // COM ProgIds must have 39 or fewer chars and no punctuation
        // other than '.'
        _progId = progId ? progId :
          wstring(L"XlOil.Rtd.") +
          wstring(clsidStr + 1, wcslen(clsidStr) - 2);
        std::replace(_progId.begin(), _progId.end(), L'-', L'.');

        HRESULT res;
        res = CoRegisterClassObject(
          clsid,                     // the CLSID to register
          _factory.p,                // the factory that can construct the object
          CLSCTX_INPROC_SERVER,      // can only be used inside our process
          REGCLS_MULTIPLEUSE,        // it can be created multiple times
          &_comRegistrationCookie);

        _regKeysAdded.push_back(
          wstring(L"Software\\Classes\\") + _progId + L"\\CLSID");
        regWriteString(
          HKEY_CURRENT_USER,
          _regKeysAdded.back().c_str(),
          0,
          clsidStr);

        // This registry entry is not needed to call CLSIDFromProgID, nor
        // to call CoCreateInstance, but for some reason the RTD call to
        // Excel will fail without it.
        _regKeysAdded.push_back(
          wstring(L"Software\\Classes\\CLSID\\") + clsidStr + L"\\InProcServer32");
        regWriteString(
          HKEY_CURRENT_USER,
          _regKeysAdded.back().c_str(),
          0,
          L"xlOil.dll");

        // Check all is good by looking up the CLISD from our progId
        CLSID foundClsid;
        res = CLSIDFromProgID(_progId.c_str(), &foundClsid);
        if (res != S_OK || !IsEqualCLSID(foundClsid, clsid))
          XLO_ERROR(L"Failed to register com server '{0}'", _progId);

#ifdef _DEBUG
        void* testObj;
        res = CoCreateInstance(
          clsid, NULL,
          CLSCTX_INPROC_SERVER,
          __uuidof(Excel::IRtdServer),
          &testObj);
        if (res != S_OK)
          XLO_ERROR(L"Failed to create com object '{0}'", _progId);
#endif

        CoTaskMemFree(clsidStr);
      }

      ~RtdManager()
      {
        CoRevokeClassObject(_comRegistrationCookie);
        for (auto& key : _regKeysAdded)
          RegDeleteKey(HKEY_CURRENT_USER, key.c_str());
      }

      void start(
        const shared_ptr<IRtdTopic>& topic) override
      {
        _server->addProducer(topic);
      }

      shared_ptr<const ExcelObj> subscribe(const wchar_t * topic) override
      {
        callRtd(topic);
        return _server->value(topic);
      }
      bool publish(const wchar_t * topic, ExcelObj&& value) override
      {
        _server->update(topic, make_shared<ExcelObj>(value));
        return true;
      }
      shared_ptr<const ExcelObj> 
        peek(const wchar_t* topic) override
      {
        // TODO: this is not enough to tell if the server is running.
        // either value is always present or...?
        return _server->value(topic);
      }
      bool drop(const wchar_t* topic) override
      {
        return _server->dropProducer(topic);
      }
      const wchar_t* progId() const noexcept override
      {
        return _progId.c_str();
      }

      // We don't use the value from the Rtd call, but it indicates to
      // Excel that the current UDF should be treated as an RTD function.
      // It prompts Excel to connect to the RTD server to link a topicId
      // to this topic string
      void callRtd(const wchar_t* topic) const
      {
        callExcel(msxll::xlfRtd, _progId, L"", topic);
      }
    };

    std::shared_ptr<IRtdManager> newRtdManager(
      const wchar_t* progId, const wchar_t* clsid)
    {
      return make_shared<COM::RtdManager>(progId, clsid);
    }
  }
}
