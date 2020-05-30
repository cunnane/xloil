#include "ComVariant.h"
#include "MessageQueue.h"

#include <xloil/RtdServer.h>
#include <xloil/ExcelObj.h>
#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/Log.h>

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

namespace xloil 
{
  namespace COM
  {
    // ATL needs this for some reason
    class AtlModule : public CAtlDllModuleT<AtlModule>
    {} theAtlModule;

    class RtdServerImpl;

    class RtdProducer : public IRtdNotify
    {
      std::atomic<bool> _cancelled;
      RtdServerImpl* _server;
      std::future<void> _taskFuture;
      shared_ptr<const ExcelObj> _value;
      wstring _topic;
      bool _persistent;
      std::mutex _lock;

    public:
      RtdProducer(
        const RtdTask& task, 
        const wchar_t* topic,
        RtdServerImpl* server,
        bool persistent)
        : _value()//CellError::GettingData)
        , _topic(topic)
        , _server(server)
        , _cancelled(false)
        , _persistent(persistent)
      {
        _taskFuture = task(*this);
      }

      ~RtdProducer()
      {
        cancel();
        _taskFuture.wait();
      }

      /// <summary>
      /// Returns true if the producer is no longer referenced and  
      /// can be scheduled for deletion
      /// </summary>
      bool disconnect(size_t remainingIds)
      {
        if (!_persistent && remainingIds == 0)
        {
          cancel();
          return true;
        }
        return false;
      }
      void cancel()
      {
        _cancelled = true;
      }
      bool complete() const
      {
        return !_taskFuture.valid()
          || _taskFuture.wait_for(std::chrono::seconds(0)) == future_status::ready;
      }
      const shared_ptr<const ExcelObj>& value()
      {
        // If the future completed, it may have set an exception so we
        // try to get the value.  After get() the future will be invalid
        // so complete() will be false, hence we only do this once.
        if (_taskFuture.valid() && complete())
        {
          try
          {
            _taskFuture.get();
          } 
          catch (const std::future_error& e)
          {
            _value.reset(new ExcelObj(e.what()));
          }
        }
        scoped_lock(_lock);
        return _value;
      }

      const wstring& topic() const
      {
        return _topic;
      }

      // IRtdNotify interface
      void publish(ExcelObj&& value) override;
      bool isCancelled() const override
      {
        return _cancelled;
      }
    };


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
        for (auto& prod : _theProducers)
          prod.second->cancel();
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
        
        std::scoped_lock lock(_lockProducers);

        _updateCallback = callback;
        _readyTopicIds.clear();

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

        std::scoped_lock lock(_lockSubscribers);

        long index[] = { 0, 0 };
        VARIANT topic;
        SafeArrayGetElement(*strings, index, &topic);

        _activeTopicIds[topicId] = topic.bstrVal;
        _subscribers[topic.bstrVal].insert(topicId);

        return S_OK;
      }

      HRESULT __stdcall raw_RefreshData(
        long* topicCount,
        SAFEARRAY** data) override
      {
        std::scoped_lock lock(_lockSubscribers);

        const auto nReady = _readyTopicIds.size();
        SAFEARRAYBOUND bounds[] = { { 2u, 0 }, { (ULONG)nReady, 0 } };
        *data = SafeArrayCreate(VT_VARIANT, 2, bounds);

        void* element;
        auto iRow = 0;
        for (auto topic : _readyTopicIds)
        {
          long index[] = { 0, iRow };
          assert(S_OK == SafeArrayPtrOfIndex(*data, index, &element));
          *(VARIANT*)element = _variant_t(topic);

          index[0] = 1;
          assert(S_OK == SafeArrayPtrOfIndex(*data, index, &element));
          *(VARIANT*)element = _variant_t();

          ++iRow;
        }

        _readyTopicIds.clear();
        *topicCount = (long)nReady;
        return S_OK;
      }

      HRESULT __stdcall raw_DisconnectData(long topicId) override
      {
        std::scoped_lock lock(_lockProducers, _lockSubscribers);

        // Reap any producers which have safely cancelled and completed
        // TODO: could do this in separate thread
        for (auto& p = _cancelledProducers.begin(); 
          p != _cancelledProducers.end();)
        {
          if ((*p)->complete())
            _cancelledProducers.erase(p++);
          else ++p;
        }

        const auto& topic = _activeTopicIds[topicId];
        if (topic.empty())
          XLO_THROW("Bad");

        auto& subscribers = _subscribers[topic];
        subscribers.erase(topicId);

        // If the disconnect() causes the producer to cancel its task,
        // it will return true here. We may not be able to just delete it, 
        // we have to wait until any threads it created exit
        auto producer = _theProducers.find(topic);
        if (producer != _theProducers.end() 
          && producer->second->disconnect(subscribers.size()))
        {
          if (!producer->second->complete())
            _cancelledProducers.push_back(producer->second);
          _theProducers.erase(topic);
        }
        
        if (subscribers.empty())
          _subscribers.erase(topic);
        _activeTopicIds.erase(topicId);

        return S_OK;
      }

      // Does nothing useful
      HRESULT __stdcall raw_Heartbeat(long* result) override
      {
        if (!result) return E_POINTER;
        *result = 1;
        return S_OK;
      }

      HRESULT __stdcall raw_ServerTerminate() override
      {
        std::scoped_lock lock(_lockProducers, _lockSubscribers);

        _activeTopicIds.clear();
        _readyTopicIds.clear();
        _updateCallback = nullptr;

        return S_OK;
      }

    private:
      std::map<wstring, shared_ptr<RtdProducer>> _theProducers;
      std::unordered_map<long, wstring> _activeTopicIds;
      std::unordered_map<wstring, unordered_set<long>> _subscribers;

      unordered_set<long> _readyTopicIds;
      std::list<shared_ptr<RtdProducer>> _cancelledProducers;

      Excel::IRTDUpdateEvent* _updateCallback;
      std::mutex _lockProducers;
      std::mutex _lockSubscribers;

    public:
      void update(const wchar_t* topic)
      {
        std::scoped_lock lock(_lockSubscribers);
        const bool nothingReady = _readyTopicIds.empty();

        const auto& subscribers = _subscribers[topic];
        _readyTopicIds.insert(subscribers.begin(), subscribers.end());

        // We only need to notify Excel about new data once. Excel will
        // only callback RefreshData approximately every 2 seconds
        if (nothingReady && _updateCallback)
        {
          queueWindowMessage([this]() { this->callUpdateNotify(); });
        }
      }

      void callUpdateNotify()
      {
        std::scoped_lock lock(_lockSubscribers);
        if (_updateCallback)
          _updateCallback->UpdateNotify();
      }

      void addProducer(const shared_ptr<RtdProducer>& job)
      {
        std::scoped_lock lock(_lockProducers);
        // TODO: can we make map key wchar_t*??
        auto iProd = _theProducers.find(job->topic());
        if (iProd != _theProducers.end())
        {
          iProd->second->cancel();
          _cancelledProducers.push_back(iProd->second);
          iProd->second = job;
        }
        else
          _theProducers.emplace(job->topic(), job);
      }

      RtdProducer* findProducer(const wchar_t* topic)
      {
        std::scoped_lock lock(_lockProducers);
        auto iJob = _theProducers.find(topic);
        if (iJob == _theProducers.end())
          return nullptr;
        return iJob->second.get();
      }
    };

    void RtdProducer::publish(ExcelObj&& value)
    {
      if (!_cancelled)
      {
        scoped_lock(_lock);
        _value.reset(new ExcelObj(std::forward<ExcelObj>(value)));
        _server->update(_topic.c_str());
      }
    }

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
        const RtdTask& task, 
        const wchar_t* topic,
        bool persistent) override
      {
        auto job = make_shared<RtdProducer>(task, topic, _server, persistent);
        _server->addProducer(job);
      }

      shared_ptr<const ExcelObj> subscribe(const wchar_t * topic) override
      {
        auto job = _server->findProducer(topic);
        callRtd(topic);
        return job ? job->value() : nullptr;
      }
      bool publish(const wchar_t * topic, ExcelObj&& value) override
      {
        auto job = _server->findProducer(topic);
        if (job)
          job->publish(std::forward<ExcelObj>(value));
        return job;
      }
      shared_ptr<const ExcelObj> peek(const wchar_t* topic) override
      {
        auto job = _server->findProducer(topic);
        return job ? job->value() : nullptr;
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
  }

  std::shared_ptr<IRtdManager> newRtdManager(
    const wchar_t* progId, const wchar_t* clsid)
  {
    return make_shared<COM::RtdManager>(progId, clsid);
  }
     
  RtdConnection::RtdConnection(IRtdManager& mgr, wstring&& topic)
    : _topic(topic)
    , _mgr(mgr)
  {
    _value = _mgr.subscribe(_topic.c_str()).get();
  }
  bool RtdConnection::hasValue() const
  {
    return !!_value;
  }
  const ExcelObj& RtdConnection::value()
  {
    return _value ? *_value : Const::Error(CellError::Null);
  }
  const ExcelObj& RtdConnection::start(const RtdTask& task)
  {
    _mgr.start(task, _topic.c_str());
    _value = _mgr.peek(_topic.c_str()).get();
    return _value ? *_value : Const::Error(CellError::GettingData);
  }

  IRtdManager* getCoreRtdServer()
  {
    // TODO: I guess we should create a mutux here, although calling
    // RTD functions from a multithreaded function is not likely to 
    // end well. Can we check for that in a non-expensive way?
    static shared_ptr<IRtdManager> ptr = newRtdManager();
    static auto deleter = Event::AutoClose() += [&]() { ptr.reset(); };
    return ptr.get();
  }

  XLOIL_EXPORT RtdConnection rtdConnect(
    IRtdManager* mgr, 
    const wchar_t* topic)
  {
    if (!mgr)
      mgr = getCoreRtdServer();
    if (!topic)
    {
      auto caller = CallerInfo().writeAddress(false);
      return RtdConnection(*mgr, std::move(caller));
    }
    return RtdConnection(*mgr, wstring(topic));
  }
}
