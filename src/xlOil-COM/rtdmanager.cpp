#include <xlOil/ExcelTypeLib.h>
#include "ComVariant.h"
#include "ClassFactory.h"
#include "RtdServerWorker.h"

#include <xlOil/ExcelThread.h>

#include <xloil/RtdServer.h>
#include <xloil/ExcelObj.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/Log.h>

#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <Objbase.h>

#include <memory>
#include <atomic>

using std::vector;
using std::shared_ptr;
using std::make_shared;
using std::wstring;
using std::atomic;
using std::mutex;

namespace
{
  // ATL needs this for some reason
  class AtlModule : public CAtlDllModuleT<AtlModule>
  {} theAtlModule;
}

namespace xloil 
{
  namespace COM
  {
    template<class TWorker>
    class
      __declspec(novtable)
      RtdServerImpl :
        public CComObjectRootEx<CComSingleThreadModel>,
        public CComCoClass<RtdServerImpl<TWorker>>,
        public IDispatchImpl<
          Excel::IRtdServer,
          &__uuidof(Excel::IRtdServer),
          &LIBID_Excel>
    {
    private:
      TWorker _worker;
      atomic<Excel::IRTDUpdateEvent*> _updateCallback;

      void updateNotify()
      {
        runExcelThread([this]()
        {
          auto callback = _updateCallback.load();
          if (callback)
            callback->raw_UpdateNotify(); // Does this really need the COM API?
        }, ExcelRunQueue::COM_API | ExcelRunQueue::ENQUEUE, 0, 1000);
      }

    public:
      TWorker& manager()
      {
        return _worker;
      }

      ~RtdServerImpl()
      {
        try
        {
          _worker.join();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RtdServer destructor: {0}", e.what());
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
        try
        {
          _updateCallback = callback;
          _worker.start([this]() { this->updateNotify(); });
          *result = 1;
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd server start: {0}", e.what());
        }
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
          _worker.connect(topicId, wstring(topicAsVariant.bstrVal));
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
        auto updates = _worker.getUpdates();
        SafeArrayGetUBound(updates, 2, topicCount);
        *data = updates;
        return S_OK;
      }

      HRESULT __stdcall raw_DisconnectData(long topicId) override
      {
        try
        {
          _worker.disconnect(topicId);
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
          // Clearing _updateCallback prevents the sending of further updates 
          // which might crash Excel if sent during teardown.
          _updateCallback = nullptr;
          _worker.quit();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Rtd Terminate: {0}", e.what());
        }
        return S_OK;
      }
    };

    class RtdServer : public IRtdServer
    {
      using ComImplType = CComObject<RtdServerImpl<RtdServerThreadedWorker<ExcelObj>>>;
      RegisterCom<ComImplType> _registrar;

      auto& server() const { return _registrar.server()->manager(); }

    public:
      RtdServer(const wchar_t* progId, const wchar_t* fixedClsid)
        : _registrar(
          [](const wchar_t*, const GUID&) { return new ComImplType(); },
          progId, 
          fixedClsid)
      {
        // TODO: why doesn't this work?
        //_registrar.cleanRegistry();
      }

      ~RtdServer()
      {
        server().join();
      }

      void start(
        const shared_ptr<IRtdPublisher>& topic) override
      {
        server().addPublisher(topic);
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
      bool publish(const wchar_t* topic, ExcelObj&& value) override
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
        return server().dropPublisher(topic);
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
          server().quit();
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