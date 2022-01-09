#pragma once
#include <Objbase.h>
#include <atlcomcli.h>
#include <xloil/Throw.h>
#include <xlOilHelpers/Environment.h>
#include <string>
#include <list>

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Simple thread-safe COM-style reference count which statisfies IUnknown.
    /// </summary>
    /// <typeparam name="Base"></typeparam>
    template <class Base>
    class ComObject : public Base
    {
    public:
      template<class...Args>
      ComObject(Args&&...args)
        : Base(std::forward<Args>(args)...)
        , m_dwRef(0)
      {}

      virtual ~ComObject()
      {
#ifdef _DEBUG
        if (m_dwRef > 0)
          _ASSERTE(0 && "Destructor called on object with positive ref count");
#endif
        // Set refcount to -(LONG_MAX/2) to protect destruction and
        // also catch mismatched Release in debug builds
        this->m_dwRef = -(LONG_MAX / 2);
      }
      ULONG STDMETHODCALLTYPE AddRef() noexcept
      {
        return ::InterlockedIncrement(&m_dwRef);
      }
      ULONG STDMETHODCALLTYPE Release() noexcept
      {
        ::InterlockedDecrement(&m_dwRef);
#ifdef _DEBUG
        if (m_dwRef < -(LONG_MAX / 2))
          _ASSERTE(0 && "Release called on a pointer that has already been released");
#endif
        if (m_dwRef == 0)
          delete this;
        return (ULONG)m_dwRef;
      }

    private:
      long m_dwRef;
    };

    class ClassFactory : public ComObject<IClassFactory>
    {
    public:
      IUnknown* _instance;

      ClassFactory(IUnknown* p)
        : _instance(p)
      {}

      STDMETHOD(CreateInstance)(
        IUnknown *pUnkOuter,
        REFIID riid,
        void **ppvObject) override
      {
        if (pUnkOuter)
          return CLASS_E_NOAGGREGATION;
        auto ret = _instance->QueryInterface(riid, ppvObject);
        return ret;
      }

      STDMETHOD(LockServer)(BOOL /*fLock*/) override
      {
        return E_NOTIMPL;
      }

      STDMETHOD(QueryInterface)(REFIID riid, void** ppv)
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
      return detail::regWriteImpl<REG_DWORD>(hive, path, name, 
        (BYTE*)&value, sizeof(DWORD));
    }

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
  
        // Ensure we populate the clsid GUID and the assoicated _clsid string. If 
        // this value has not been provided or read from the registry above, we
        // create a new GUID.

        if (fixedClsid)
          _clsid = fixedClsid;
         
        if (!_clsid.empty())
        {
           CLSIDFromString(_clsid.c_str(), &clsid);
        }
        else
        {
          auto fail = CoCreateGuid(&clsid) != 0;

          // This generates the string '{W-X-Y-Z}'
          wchar_t clsidStr[128];
          fail = fail || StringFromGUID2(clsid, clsidStr, _countof(clsidStr)) == 0;
          if (fail)
            XLO_THROW("Failed to create CLSID for COM Server");

          _clsid = clsidStr;
        }

        // If no progId has been specified, use 'XlOil.<Clsid>'
        if (!progId)
        {
          // COM ProgIds must have 39 or fewer chars and no punctuation other than '.'
          _progId = std::wstring(L"XlOil.") + _clsid.substr(1, _clsid.size() - 2);
          std::replace(_progId.begin(), _progId.end(), L'-', L'.');
        }

        // Create the COM 'server' and a class factory which returns it.  Normally
        // a class factory creates the COM object on demand, so we are subverting
        // the pattern slightly!
        _server  = createServer(_progId.c_str(), clsid);
        _factory = new ClassFactory((IDispatch*)_server.p);

        // Register our class factory in the Registry
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
      const wchar_t* clsid()  const { return _clsid.c_str(); }
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
  
    /// <summary>
    /// Null implementation of all IDispatch functionality. Returns E_NOTIMPL
    /// for every call.
    /// </summary>
    /// <typeparam name="TBase"></typeparam>
    template <class TBase>
    class NoIDispatchImpl : public TBase
    {
    public:
      // IDispatch
      STDMETHOD(GetTypeInfoCount)(UINT* /*pctinfo*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(GetTypeInfo)(
        UINT /*itinfo*/,
        LCID /*lcid*/,
        ITypeInfo** /*pptinfo*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(GetIDsOfNames)(
        REFIID /*riid*/,
        LPOLESTR* /*rgszNames*/,
        UINT /*cNames*/,
        LCID /*lcid*/,
        DISPID* /*rgdispid*/)
      {
        return E_NOTIMPL;
      }
      STDMETHOD(Invoke)(
        DISPID /*dispidMember*/,
        REFIID /*riid*/,
        LCID /*lcid*/,
        WORD /*wFlags*/,
        DISPPARAMS* /*pdispparams*/,
        VARIANT* /*pvarResult*/,
        EXCEPINFO* /*pexcepinfo*/,
        UINT* /*puArgErr*/)
      {
        return E_NOTIMPL;
      }
    };

    /// <summary>
    /// Base class for a COM event sink which binds to a single source and responds
    /// to QueryInterface for the TSource interface.
    /// </summary>
    /// <typeparam name="TBase">Base class, must inherit from TSource</typeparam>
    /// <typeparam name="TSource">Interface for target connection point</typeparam>
    template<class TBase, class TSource>
    class ComEventHandler : public TBase
    {
    public:
      ComEventHandler() noexcept 
        : _pIConnectionPoint(nullptr)
      {}

      virtual ~ComEventHandler() noexcept
      {
        // We do not disconnect as should never enter the dtor while connected
      }

      bool connect(IUnknown* source) noexcept
      {
        IConnectionPointContainer* pContainer = nullptr;

        // Get connection point for source
        source->QueryInterface(IID_IConnectionPointContainer, (void**)&pContainer);
        if (pContainer)
        {
          pContainer->FindConnectionPoint(__uuidof(TSource), &_pIConnectionPoint);
          pContainer->Release();
        }

        return _pIConnectionPoint && (S_OK == _pIConnectionPoint->Advise(this, &_dwEventCookie));
      }

      /// <summary>
      /// Must be called to destroy this object because the connection creats 
      /// a reference to this class, preventing it from being destroyed.
      /// </summary>
      void disconnect() noexcept
      {
        if (_pIConnectionPoint)
        {
          _pIConnectionPoint->Unadvise(_dwEventCookie);
          _dwEventCookie = 0;
          _pIConnectionPoint->Release();
          _pIConnectionPoint = NULL;
        }
      }

      STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject) noexcept
      {
        if (riid == IID_IUnknown)
          *ppvObject = (IUnknown*)this;
        else if ((riid == IID_IDispatch) || (riid == __uuidof(TSource)))
          *ppvObject = (IDispatch*)this;
        else
          return E_NOINTERFACE;

        AddRef();
        return S_OK;
      }

    private:
      IConnectionPoint* _pIConnectionPoint;
      DWORD	_dwEventCookie;
    };
  }
}