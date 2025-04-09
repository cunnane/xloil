#pragma once
#include <xloil/WindowsSlim.h>
#include <Objbase.h>
#include <atlcomcli.h>
#include <xloil/Throw.h>
#include <xloil/Log.h>
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
      ULONG __stdcall AddRef() noexcept
      {
        return ::InterlockedIncrement(&m_dwRef);
      }
      ULONG __stdcall Release() noexcept
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

    class RegisterCom
    {
      CComPtr<IUnknown> _factory;
      DWORD _comRegistrationCookie;
      wchar_t _clsidStr[39]; // includes braces and null-terminator
      std::wstring _progId;
      std::list<std::pair<HKEY, std::wstring>> _regKeysAdded;

    public:
      RegisterCom(
        const std::function<IUnknown* ()>& createServer,
        const wchar_t* progId = nullptr,
        const GUID* fixedClsid = nullptr);

      ~RegisterCom();

      const wchar_t* progid() const { return _progId.c_str(); }
      const wchar_t* clsid()  const { return _clsidStr; }

      HRESULT writeRegistry(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        const wchar_t* value);

      HRESULT writeRegistry(
        HKEY hive,
        const wchar_t* path,
        const wchar_t* name,
        DWORD value);

      void addedKey(HKEY hive, std::wstring_view path)
      {
        _regKeysAdded.emplace_back(std::pair(hive, std::wstring(path)));
      }

      void cleanRegistry();
    };
  
    /// <summary>
    /// Null implementation of all IDispatch functionality. Returns E_NOTIMPL
    /// for every call.
    /// </summary>
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
        if (_pIConnectionPoint)
          XLO_ERROR("ComEventHandler destroyed before being disconnected");
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
      /// Must be called to destroy this object because the connection creates 
      /// a reference to this class, preventing it from being destroyed.
      /// </summary>
      void disconnect() noexcept
      {
        if (_pIConnectionPoint)
        {
          _pIConnectionPoint->Unadvise(_dwEventCookie);
          _pIConnectionPoint->Release();
          _pIConnectionPoint = nullptr;
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