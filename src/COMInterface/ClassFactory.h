#pragma once
#include <list>
#include <atlbase.h>
#include <atlcom.h>
#include <atlwin.h>
#include <Objbase.h>
#include <string>
#include <xloil/Throw.h>

namespace xloil
{
  namespace COM
  {
    template <class TInstance>
    class ClassFactory : public IClassFactory
    {
    public:
      TInstance* _instance;

      ClassFactory(TInstance* p)
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

    HRESULT inline regWriteString(
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

    template <class TInstance>
    class RegisterCom
    {
      CComPtr<TInstance> _server;
      CComPtr<ClassFactory<TInstance>> _factory;
      DWORD _comRegistrationCookie;
      std::wstring _clsid;
      std::wstring _progId;
      std::list<std::wstring> _regKeysAdded;

    public:
      RegisterCom(const wchar_t* progId, const wchar_t* fixedClsid)
      {
        _server = new CComObject<TInstance>();
        _factory = new ClassFactory<TInstance>(_server.p);

        if (progId && !fixedClsid)
          XLO_THROW("If you specify a ProgId you must also specify a "
            "CLSID or different Excel instances may clash");

        GUID clsid;
        HRESULT hCreateGuid = fixedClsid
          ? CLSIDFromString(fixedClsid, &clsid)
          : CoCreateGuid(&clsid);

        LPOLESTR clsidStr;
        // This generates the string '{W-X-Y-Z}'
        StringFromCLSID(clsid, &clsidStr);
        _clsid = clsidStr;
        CoTaskMemFree(clsidStr);

        using std::wstring;

        // COM ProgIds must have 39 or fewer chars and no punctuation
        // other than '.'
        _progId = progId ? progId :
          wstring(L"XlOil.") + _clsid.substr(1, _clsid.size() - 2);
        std::replace(_progId.begin(), _progId.end(), L'-', L'.');

        HRESULT res;
        res = CoRegisterClassObject(
          clsid,                     // the CLSID to register
          _factory.p,                // the factory that can construct the object
          CLSCTX_INPROC_SERVER,      // can only be used inside our process
          REGCLS_MULTIPLEUSE,        // it can be created multiple times
          &_comRegistrationCookie);

        writeRegistry(
          HKEY_CURRENT_USER,
          (wstring(L"Software\\Classes\\") + _progId + L"\\CLSID").c_str(),
          0,
          _clsid.c_str());

        // This registry entry is not needed to call CLSIDFromProgID, nor
        // to call CoCreateInstance, but for some reason the RTD call to
        // Excel will fail without it.
        writeRegistry(
          HKEY_CURRENT_USER,
          (wstring(L"Software\\Classes\\CLSID\\") + _clsid + L"\\InProcServer32").c_str(),
          0,
          L"xlOil.dll");

        // Check all is good by looking up the CLISD from our progId
        CLSID foundClsid;
        res = CLSIDFromProgID(_progId.c_str(), &foundClsid);
        if (res != S_OK || !IsEqualCLSID(foundClsid, clsid))
          XLO_THROW(L"Failed to register com server '{0}'", _progId);
      }

      const wchar_t* progid() const { return _progId.c_str(); }
      const wchar_t* clsid() const { return _clsid.c_str(); }
      TInstance& server() const { return *_server; }

      HRESULT writeRegistry(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        const wchar_t* value)
      {
        _regKeysAdded.push_back(path);
        return regWriteString(
          hive,
          path,
          name,
          value);
      }

      ~RegisterCom()
      {
        CoRevokeClassObject(_comRegistrationCookie);

        // Remove the registry keys used to locate the server (they would be 
        // removed anyway on windows logoff)
        for (auto& key : _regKeysAdded)
          RegDeleteKey(HKEY_CURRENT_USER, key.c_str());
      }
    };
  
  }
}