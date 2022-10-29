#include "ClassFactory.h"
#include <xlOilHelpers/GuidUtils.h>

namespace
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


  HRESULT inline regWrite(
    HKEY hive,
    const wchar_t* path,
    const wchar_t* name,
    const wchar_t* value)
  {
    return regWriteImpl<REG_SZ>(hive, path, name,
      (BYTE*)value, (wcslen(value) + 1) * sizeof(wchar_t));
  }

  HRESULT inline regWrite(
    HKEY hive,
    const wchar_t* path,
    const wchar_t* name,
    DWORD value)
  {
    return regWriteImpl<REG_DWORD>(hive, path, name,
      (BYTE*)&value, sizeof(DWORD));
  }
}

namespace xloil
{
  namespace COM
  {
    class ClassFactory : public ComObject<IClassFactory>
    {
    public:
      using creator_t = std::function<IUnknown* ()>;
      creator_t _creator;

      ClassFactory(const creator_t& func)
        : _creator(func)
      {}

      STDMETHOD(CreateInstance)(
        IUnknown* pUnkOuter,
        REFIID riid,
        void** ppvObject) override
      {
        if (pUnkOuter)
          return CLASS_E_NOAGGREGATION;
        auto* instance = _creator();
        auto ret = instance->QueryInterface(riid, ppvObject);
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

    RegisterCom::RegisterCom(
      const std::function<IUnknown* ()>& createServer,
      const wchar_t* progId,
      const GUID* fixedClsid)
    {
      GUID clsid = { 0 };

      if (fixedClsid)
      {
        clsid = *fixedClsid;
        StringFromGUID2(clsid, _clsidStr, _countof(_clsidStr));
      }

      if (progId)
      {
        // Check if ProgId is already registered by trying to find its CLSID
        auto clsidKey = fmt::format(L"Software\\Classes\\{0}\\CLSID", progId);
        std::wstring registeredClsId;
        if (getWindowsRegistryValue(L"HKCU", clsidKey.c_str(), registeredClsId))
        {
          if (fixedClsid)
          {
            if (_wcsicmp(_clsidStr, registeredClsId.c_str()) != 0)
              XLO_THROW(L"COM Server progId={0} already in registry with clsid={1}, "
                "but clsid={2} was requested",
                progId, registeredClsId, _clsidStr);

          }
          else
          {
            wcsncpy_s(_clsidStr, registeredClsId.c_str(), registeredClsId.length());
            CLSIDFromString(_clsidStr, &clsid);
          }
        }
        _progId = progId;
      }

      // Ensure we populate the clsid GUID and the assoicated _clsid string. If 
      // this value has not been provided or read from the registry above, we
      // create a new GUID.

      if (clsid.Data1 == 0)
      {
        if (CoCreateGuid(&clsid) != 0)
          XLO_THROW("Failed to create CLSID for COM Server");
        StringFromGUID2(clsid, _clsidStr, _countof(_clsidStr));
      }

      // If no progId has been specified, use 'XlOil.<Clsid>'
      if (!progId)
      {
        // COM ProgIds must have 39 or fewer chars and no punctuation other than '.'
        _progId = std::wstring(L"XlOil.") + guidToWString(clsid, GuidToString::BASE62);
      }

      // Create the COM 'server' and a class factory which returns it.  Normally
      // a class factory creates the COM object on demand, so we are subverting
      // the pattern slightly!
      _factory = new ClassFactory(createServer);

      // Register our class factory in the Registry
      HRESULT res;
      res = CoRegisterClassObject(
        clsid,                     // the CLSID to register
        _factory.p,                // the factory that can construct the object
        CLSCTX_INPROC_SERVER,      // can only be used inside our process
        REGCLS_MULTIPLEUSE,        // it can be created multiple times
        &_comRegistrationCookie);

      auto keyPath = fmt::format(L"Software\\Classes\\{0}\\CLSID", _progId);
      writeRegistry(
        HKEY_CURRENT_USER,
        keyPath.c_str(),
        0,
        _clsidStr);

      // Note the outer key to ensure it is deleted
      addedKey(HKEY_CURRENT_USER,
        std::wstring_view(keyPath).substr(0, keyPath.find_last_of(L'\\')));

      // This registry entry is not needed to call CLSIDFromProgID, nor
      // to call CoCreateInstance, but for some reason the RTD call to
      // Excel will fail without it.
      keyPath = fmt::format(L"Software\\Classes\\CLSID\\{0}\\InProcServer32", _clsidStr);
      writeRegistry(
        HKEY_CURRENT_USER,
        keyPath.c_str(),
        0,
        L"xlOil.dll"); // Name of dll isn't actually used.

      addedKey(HKEY_CURRENT_USER,
        std::wstring_view(keyPath).substr(0, keyPath.find_last_of(L'\\')));

      // Check all is good by looking up the CLSID from our progId
      CLSID foundClsid;
      res = CLSIDFromProgID(_progId.c_str(), &foundClsid);
      if (res != S_OK || !IsEqualCLSID(foundClsid, clsid))
        XLO_THROW(L"Failed to register com server '{0}'", _progId);
    }

    HRESULT RegisterCom::writeRegistry(
      HKEY hive,
      const wchar_t* path,
      const wchar_t* name,
      const wchar_t* value)
    {
      addedKey(hive, path);
      XLO_TRACE(L"Writing registry key {}\\{} = {}", path, name ? name : L"", value);
      return regWrite(hive, path, name, value);
    }

    HRESULT RegisterCom::writeRegistry(
      HKEY hive,
      const wchar_t* path,
      const wchar_t* name,
      DWORD value)
    {
      addedKey(hive, path);
      XLO_TRACE(L"Writing registry key {}\\{} = {}", path, name, value);
      return regWrite(hive, path, name, value);
    }

    void RegisterCom::cleanRegistry()
    {
      // Remove the registry keys used to locate the server (they would be 
      // removed anyway on windows logoff)
      for (auto& [hive, path] : _regKeysAdded)
        RegDeleteKey(hive, path.c_str());
      _regKeysAdded.clear();
    }

    RegisterCom::~RegisterCom()
    {
      CoRevokeClassObject(_comRegistrationCookie);
      cleanRegistry();
    }
  }
}
