#pragma once
#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <Objbase.h>
#include <xloil/Throw.h>
#include <xlOilHelpers/Environment.h>
#include <string>
#include <list>

namespace xloil
{
  namespace COM
  {
    class ClassFactory : public IClassFactory
    {
    public:
      IUnknown* _instance;

      ClassFactory(IUnknown* p)
        : _instance(p)
        , _cRef(0)
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

      HRESULT STDMETHODCALLTYPE LockServer(BOOL /*fLock*/) override
      {
        return E_NOTIMPL;
      }

      HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv)
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

    namespace detail
    {
      template<int TValCode>
      HRESULT inline regWriteImpl(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        BYTE* data,
        size_t dataLength)
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
          TValCode,
          data,
          (DWORD)dataLength);

        RegCloseKey(key);
        return res;
      }
    }

    HRESULT inline regWrite(
      HKEY hive,
      const wchar_t* path,
      const wchar_t* name,
      const wchar_t* value)
    {
      return detail::regWriteImpl<REG_SZ>(hive, path, name, 
        (BYTE*)value, (wcslen(value) + 1) * sizeof(wchar_t));
    }
    HRESULT inline regWrite(
      HKEY hive,
      const wchar_t* path,
      const wchar_t* name,
      DWORD value)
    {
      return detail::regWriteImpl<REG_DWORD>(hive, path, name, (BYTE*)&value, sizeof(DWORD));
    }

    /// <summary>
    /// A copy of CComObject from ATL but with a modern forwarding constructor
    /// </summary>
    /// <typeparam name="Base"></typeparam>
    template <class Base>
    class ComObject : public Base
    {
    public:
      typedef Base _BaseClass;

      template<class...Args>
      ComObject(Args&&...args)
        : Base(std::forward<Args>(args)...)
      {
        _pAtlModule->Lock();
      }

      virtual ~ComObject()
      {
        // Set refcount to -(LONG_MAX/2) to protect destruction and
        // also catch mismatched Release in debug builds
        this->m_dwRef = -(LONG_MAX / 2);
        this->FinalRelease();
        _pAtlModule->Unlock();
      }
      STDMETHOD_(ULONG, AddRef)()
      {
        return this->InternalAddRef();
      }
      STDMETHOD_(ULONG, Release)()
      {
        ULONG l = this->InternalRelease();
        if (l == 0)
        {
          // Lock the module to avoid DLL unload when destruction of member variables take a long time
          ModuleLockHelper lock;
          delete this;
        }
        return l;
      }
      STDMETHOD(QueryInterface)(
        REFIID iid,
        _COM_Outptr_ void** ppvObject) throw()
      {
        return this->_InternalQueryInterface(iid, ppvObject);
      }
      // TODO: doesnt this exist in IUnk?
      /*template <class Q>
      HRESULT STDMETHODCALLTYPE QueryInterface(
        _COM_Outptr_ Q** pp) throw()
      {
        return QueryInterface(__uuidof(Q), (void**)pp);
      }*/
    };

    template <class TComServer>
    class RegisterCom
    {
      CComPtr<TComServer> _server;
      CComPtr<ClassFactory> _factory;
      DWORD _comRegistrationCookie;
      std::wstring _clsid;
      std::wstring _progId;
      std::list<std::wstring> _regKeysAdded;

    public:
      template<class TCreatorFunc>
      RegisterCom(
        TCreatorFunc createServer,
        const wchar_t* progId = nullptr,
        const wchar_t* fixedClsid = nullptr)
      {
        GUID clsid;

        if (progId)
        {
          // Check if ProgId is already registered by trying to find its CLSID
          auto clsidKey = fmt::format(L"Software\\Classes\\{0}\\CLSID", progId);
          if (getWindowsRegistryValue(L"HKCU", clsidKey.c_str(), _clsid))
          {
            if (fixedClsid && _wcsicmp(_clsid.c_str(), fixedClsid) != 0)
              XLO_THROW(L"COM Server progId={0} already in registry with clsid={1}, "
                         "but clsid={2} was requested",
                          progId, _clsid, fixedClsid);
            
          }
          _progId = progId;
        }
  

        // If no CLSID, generate one
        if (fixedClsid)
          _clsid = fixedClsid;
         
        
        if (!_clsid.empty())
        {
           CLSIDFromString(_clsid.c_str(), &clsid);
        }
        else
        {
          auto fail = CoCreateGuid(&clsid) != 0;
          LPOLESTR clsidStr = nullptr;
          // This generates the string '{W-X-Y-Z}'
          fail = fail || StringFromCLSID(clsid, &clsidStr) != 0;
          if (fail)
            XLO_THROW("Failed to create CLSID for COM Server");
          _clsid = clsidStr;
          CoTaskMemFree(clsidStr);
        }

        if (!progId)
        {
          // COM ProgIds must have 39 or fewer chars and no punctuation other than '.'
          _progId = std::wstring(L"XlOil.") + _clsid.substr(1, _clsid.size() - 2);
          std::replace(_progId.begin(), _progId.end(), L'-', L'.');
        }

        _server  = createServer(_progId.c_str(), clsid);
        _factory = new ClassFactory((IDispatch*)_server.p);

        HRESULT res;
        res = CoRegisterClassObject(
          clsid,                     // the CLSID to register
          _factory.p,                // the factory that can construct the object
          CLSCTX_INPROC_SERVER,      // can only be used inside our process
          REGCLS_MULTIPLEUSE,        // it can be created multiple times
          &_comRegistrationCookie);

        writeRegistry(
          HKEY_CURRENT_USER,
          fmt::format(L"Software\\Classes\\{0}\\CLSID", _progId).c_str(),
          0,
          _clsid.c_str());

        // Add to our list of added keys, to ensure outer key is deleted
        _regKeysAdded.emplace_back(fmt::format(L"Software\\Classes\\{0}", _progId));

        // This registry entry is not needed to call CLSIDFromProgID, nor
        // to call CoCreateInstance, but for some reason the RTD call to
        // Excel will fail without it.
        writeRegistry(
          HKEY_CURRENT_USER,
          fmt::format(L"Software\\Classes\\CLSID\\{0}\\InProcServer32", _clsid).c_str(),
          0,
          L"xlOil.dll"); // Name of dll isn't actually used.
        _regKeysAdded.emplace_back(fmt::format(L"Software\\Classes\\CLSID\\{0}", _clsid));

        // Check all is good by looking up the CLISD from our progId
        CLSID foundClsid;
        res = CLSIDFromProgID(_progId.c_str(), &foundClsid);
        if (res != S_OK || !IsEqualCLSID(foundClsid, clsid))
          XLO_THROW(L"Failed to register com server '{0}'", _progId);
      }

      const wchar_t* progid() const { return _progId.c_str(); }
      const wchar_t* clsid() const { return _clsid.c_str(); }
      const CComPtr<TComServer>& server() const { return _server; }

      HRESULT writeRegistry(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        const wchar_t* value)
      {
        _regKeysAdded.push_back(path);
        return regWrite(hive, path, name, value);
      }
      HRESULT writeRegistry(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        DWORD value)
      {
        _regKeysAdded.push_back(path);
        return regWrite(hive, path, name, value);
      }
      void cleanRegistry()
      {
        // Remove the registry keys used to locate the server (they would be 
        // removed anyway on windows logoff)
        // TODO: only removes from HKCU
        for (auto& key : _regKeysAdded)
          RegDeleteKey(HKEY_CURRENT_USER, key.c_str());
        _regKeysAdded.clear();
      }
      ~RegisterCom()
      {
        CoRevokeClassObject(_comRegistrationCookie);
        cleanRegistry();
      }
    };
  
    template <class T>
      class __declspec(novtable) NoIDispatchImpl :
      public T
    {
    public:
      // IDispatch
      STDMETHOD(GetTypeInfoCount)(_Out_ UINT* /*pctinfo*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(GetTypeInfo)(
        UINT /*itinfo*/,
        LCID /*lcid*/,
        _Outptr_result_maybenull_ ITypeInfo** /*pptinfo*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(GetIDsOfNames)(
        _In_ REFIID /*riid*/,
        _In_reads_(cNames) _Deref_pre_z_ LPOLESTR* /*rgszNames*/,
        _In_range_(0, 16384) UINT /*cNames*/,
        LCID /*lcid*/,
        _Out_ DISPID* /*rgdispid*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(Invoke)(
        _In_ DISPID /*dispidMember*/,
        _In_ REFIID /*riid*/,
        _In_ LCID /*lcid*/,
        _In_ WORD /*wFlags*/,
        _In_ DISPPARAMS* /*pdispparams*/,
        _Out_opt_ VARIANT* /*pvarResult*/,
        _Out_opt_ EXCEPINFO* /*pexcepinfo*/,
        _Out_opt_ UINT* /*puArgErr*/)
      {
        return E_NOTIMPL;
      }
    };

    template<class TSource>
      void connectSourceToSink(
        const IID& iid,
        TSource* source,
        IDispatch* sink,
        IConnectionPoint*& connectionPoint,
        DWORD& eventCookie)
      {
        IConnectionPointContainer* pContainer;
        IUnknown* pIUnknown = nullptr;

        // Get IUnknown for sink
        sink->QueryInterface(IID_IUnknown, (void**)(&pIUnknown));

        // Get connection point for source
        source->QueryInterface(IID_IConnectionPointContainer, (void**)&pContainer);
        if (pContainer)
        {
          pContainer->FindConnectionPoint(iid, &connectionPoint);
          pContainer->Release();
        }

        if (connectionPoint)
          connectionPoint->Advise(pIUnknown, &eventCookie);

        pIUnknown->Release();
    }
  }
}