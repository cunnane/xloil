
#include "ClassFactory.h"
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/Ribbon.h>
#include <xloil/Throw.h>
#include <xloil/Log.h>
#include <atlctl.h>

namespace xloil
{
  namespace COM
  {
    class __declspec(novtable)
      CustomTaskPaneConsumerImpl :
      public CComObjectRootEx<CComSingleThreadModel>,
      public NoIDispatchImpl<Office::ICustomTaskPaneConsumer>
    {
    public:

      CustomTaskPaneConsumerImpl()
      {
      }
      ~CustomTaskPaneConsumerImpl()
      {}

      virtual HRESULT __stdcall raw_CTPFactoryAvailable(
        Office::ICTPFactory* factory
      ) override
      {

      }

      HRESULT _InternalQueryInterface(REFIID riid, void** ppv) throw()
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == IID_IDispatch
          || riid == __uuidof(ICustomTaskPaneConsumer))
        {
          *ppv = this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }
#pragma region IDispatch

      STDMETHOD(GetTypeInfoCount)(_Out_ UINT* /*pctinfo*/)
      {
        return 0;
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
        _In_reads_(cNames) _Deref_pre_z_ LPOLESTR* rgszNames,
        _In_range_(0, 16384) UINT cNames,
        LCID /*lcid*/,
        _Out_ DISPID* rgdispid)
      {
        return E_NOTIMPL;
      }

      STDMETHOD(Invoke)(
        _In_ DISPID dispidMember,
        _In_ REFIID /*riid*/,
        _In_ LCID /*lcid*/,
        _In_ WORD /*wFlags*/,
        _In_ DISPPARAMS* pdispparams,
        _Out_opt_ VARIANT* pvarResult,
        _Out_opt_ EXCEPINFO* /*pexcepinfo*/,
        _Out_opt_ UINT* /*puArgErr*/)
      {
        // Remember the args are in reverse order
        auto* rgvarg = pdispparams->rgvarg;

        if (dispidMember == 1)
        {
          return raw_CTPFactoryAvailable((Office::ICTPFactory*)rgvarg[0].pdispVal);
        }
        else
        {
          XLO_ERROR("Internal Error: unknown dispid called on task pane consumer Invoke.");
          return E_FAIL;
        }
        return S_OK;
      }

#pragma endregion

    };

    //public class zCustomTaskPaneCollection
    //{
    //  // Public list of TaskPane items
    //  public List<zCustomTaskPane> Items = new List<zCustomTaskPane>();
    //  private Office.ICTPFactory _paneFactory;

    //};



/////////////////////////////////////////////////////////////////////////////

  class ATL_NO_VTABLE CustomTaskPaneCtrl :
      public CComObjectRootEx<CComSingleThreadModel>,
      public IDispatchImpl<IDispatch>,
      //public CStockPropImpl<IDispatch, IPolyCtl, &IID_IPolyCtl, &LIBID_POLYGONLib>,
      public CComControl<CustomTaskPaneCtrl>,
      //public IPersistStreamInitImpl<CustomTaskPaneCtrl>,
      public IOleControlImpl<CustomTaskPaneCtrl>,
      public IOleObjectImpl<CustomTaskPaneCtrl>,
      public IOleInPlaceActiveObjectImpl<CustomTaskPaneCtrl>,
      public IViewObjectExImpl<CustomTaskPaneCtrl>,
      public IOleInPlaceObjectWindowlessImpl<CustomTaskPaneCtrl>
      //public ISupportErrorInfo,
      //public IConnectionPointContainerImpl<CustomTaskPaneCtrl>,
      //public IPersistStorageImpl<CustomTaskPaneCtrl>,
      //public ISpecifyPropertyPagesImpl<CustomTaskPaneCtrl>,
      //public IQuickActivateImpl<CustomTaskPaneCtrl>,
      //public IDataObjectImpl<CustomTaskPaneCtrl>,
      //public IProvideClassInfo2Impl<&CLSID_PolyCtl, &DIID__IPolyCtlEvents, &LIBID_POLYGONLib>,
      //public IPropertyNotifySinkCP<CustomTaskPaneCtrl>,
      //public CComCoClass<CustomTaskPaneCtrl, &CLSID_PolyCtl>,
      //public CProxy_IPolyCtlEvents< CustomTaskPaneCtrl >,
      //public IObjectSafetyImpl<CustomTaskPaneCtrl, INTERFACESAFE_FOR_UNTRUSTED_CALLER>
    {
      GUID _clsid;

    public:
      CustomTaskPaneCtrl(const wchar_t* progId, const GUID& clsid)
        : _clsid(clsid)
      {
        CComControlBase::m_bWindowOnly = true;
      }
      static const CLSID& WINAPI GetObjectCLSID()
      {
        XLO_THROW("Not supported");
      }

      HWND getHwnd() const
      {
        return CComControlBase::m_hWndCD;
      }

      //DECLARE_REGISTRY_RESOURCEID(IDR_POLYCTL)

      BEGIN_COM_MAP(CustomTaskPaneCtrl)
        //COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
        //COM_INTERFACE_ENTRY(IPolyCtl)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IViewObjectEx)
        COM_INTERFACE_ENTRY(IViewObject2)
        COM_INTERFACE_ENTRY(IViewObject)
        //COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
        COM_INTERFACE_ENTRY(IOleInPlaceObject)
        COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObject)
        //COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
        COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
        COM_INTERFACE_ENTRY(IOleControl)
        COM_INTERFACE_ENTRY(IOleObject)
        //COM_INTERFACE_ENTRY(IPersistStreamInit)
        //COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
        //COM_INTERFACE_ENTRY(ISupportErrorInfo)
        //COM_INTERFACE_ENTRY(IConnectionPointContainer)
        //COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
        //COM_INTERFACE_ENTRY(IQuickActivate)
        //COM_INTERFACE_ENTRY(IPersistStorage)
        //COM_INTERFACE_ENTRY(IDataObject)
       // COM_INTERFACE_ENTRY(IProvideClassInfo)
        //COM_INTERFACE_ENTRY(IProvideClassInfo2)
        //COM_INTERFACE_ENTRY(IObjectSafety)
      END_COM_MAP()

      //BEGIN_PROP_MAP(CustomTaskPaneCtrl)
   
      //END_PROP_MAP()

      //BEGIN_CONNECTION_POINT_MAP(CustomTaskPaneCtrl)
      ////  CONNECTION_POINT_ENTRY(DIID__IPolyCtlEvents)
      ////  CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
      //END_CONNECTION_POINT_MAP()

      BEGIN_MSG_MAP(CustomTaskPaneCtrl)
        CHAIN_MSG_MAP(CComControl<CustomTaskPaneCtrl>)
        DEFAULT_REFLECTION_HANDLER()
      END_MSG_MAP()
      // Handler prototypes:
      //  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
      //  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
      //  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
      /*
      BEGIN_MSG_MAP(CCDInfo)
	      MESSAGE_HANDLER(WM_PAINT         , OnPaint)
	      MESSAGE_HANDLER(WM_ERASEBKGND , OnEraseBkgnd)
	      MESSAGE_HANDLER(WM_MOUSEMOVE     , OnMouseMove)
	      MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
	      MESSAGE_HANDLER(WM_LBUTTONUP  , RelayEvent)
	      MESSAGE_HANDLER(WM_RBUTTONDOWN, RelayEvent)
	      MESSAGE_HANDLER(WM_RBUTTONUP  , RelayEvent)
	      MESSAGE_HANDLER(WM_MBUTTONDOWN, RelayEvent)
	      MESSAGE_HANDLER(WM_MBUTTONUP  , RelayEvent)
        END_MSG_MAP()
      */


      // ISupportsErrorInfo
      //STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
      //{
      //  static const IID* arr[] =
      //  {
      //    &IID_IPolyCtl,
      //  };
      //  for (int i = 0; i < sizeof(arr) / sizeof(arr[0]); i++)
      //  {
      //    if (InlineIsEqualGUID(*arr[i], riid))
      //      return S_OK;
      //  }
      //  return S_FALSE;
      //}

      // IViewObjectEx
      DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

      // IPolyCtl
    public:
      STDMETHOD(EnumVerbs)(_Outptr_ IEnumOLEVERB** ppEnumOleVerb)
      {
        if (!ppEnumOleVerb)
          return E_POINTER;
        return OleRegEnumVerbs(_clsid, ppEnumOleVerb);
      }
      STDMETHOD(GetUserClassID)(_Out_ CLSID* pClsid)
      {
        if (!pClsid)
          return E_POINTER;
        *pClsid = _clsid;
        return S_OK;
      }
      STDMETHOD(GetUserType)(
        _In_ DWORD dwFormOfType,
        _Outptr_result_z_ LPOLESTR* pszUserType)
      {
        return OleRegGetUserType(_clsid, dwFormOfType, pszUserType);
      }
      STDMETHOD(GetMiscStatus)(
        _In_ DWORD dwAspect,
        _Out_ DWORD* pdwStatus)
      {
        return OleRegGetMiscStatus(_clsid, dwAspect, pdwStatus);
      }

      HRESULT CustomTaskPaneCtrl::OnDraw(ATL_DRAWINFO& di)
      {
        return S_OK;
      }
    };

    class __declspec(novtable) CustomTaskPaneEventHandler
      : public CComObjectRootEx<CComSingleThreadModel>, 
        public NoIDispatchImpl<Office::_CustomTaskPaneEvents>
    {
    public:
      CustomTaskPaneEventHandler(ICustomTaskPane& parent)
        : _parent(parent)
      {
      }

      void connect(Office::_CustomTaskPane* source)
      {
        connectSourceToSink(__uuidof(Office::_CustomTaskPaneEvents),
          source, this, _pIConnectionPoint, _dwEventCookie);
      }
      virtual ~CustomTaskPaneEventHandler()
      {
        close();
      }

      void close()
      {
        if (_pIConnectionPoint)
        {
          _pIConnectionPoint->Unadvise(_dwEventCookie);
          _dwEventCookie = 0;
          _pIConnectionPoint->Release();
          _pIConnectionPoint = NULL;
        }
      }

      void addVisibilityChangeHandler(const ICustomTaskPane::ChangeHandler& handler)
      {
        _visibilityChangeHandlers.push_back(handler);
      }
      void addDockStateChangeHandler(const ICustomTaskPane::ChangeHandler& handler)
      {
        _dockStateChangeHandlers.push_back(handler);
      }

      HRESULT _InternalQueryInterface(REFIID riid, void** ppv) throw()
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == IID_IDispatch
          || riid == __uuidof(Office::_CustomTaskPaneEvents))
        {
          *ppv = this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }

      STDMETHOD(Invoke)(DISPID dispidMember, REFIID /*riid*/,
        LCID /*lcid*/, WORD /*wFlags*/, DISPPARAMS* pdispparams, VARIANT* /*pvarResult*/,
        EXCEPINFO* /*pexcepinfo*/, UINT* /*puArgErr*/)
      {
        try
        {
          auto* rgvarg = pdispparams->rgvarg;

          // These dispids are copied from oleview and are in the same order as listed there
          switch (dispidMember)
          {
          case 1:
            VisibleStateChange((_CustomTaskPane*)rgvarg[0].pdispVal);
            break;
          case 2:
            DockPositionStateChange((_CustomTaskPane*)rgvarg[0].pdispVal);
            break;
          }
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Error during COM event handler callback: {0}", e.what());
        }

        return S_OK;
      }

    private:
      HRESULT VisibleStateChange(
        struct _CustomTaskPane* CustomTaskPaneInst)
      {
        for (auto & h : _visibilityChangeHandlers)
          h(_parent);
        return S_OK;
      }
      HRESULT DockPositionStateChange(
        struct _CustomTaskPane* CustomTaskPaneInst)
      {
        for (auto & h : _visibilityChangeHandlers)
          h(_parent);
        return S_OK;
      }

    private:
      IConnectionPoint* _pIConnectionPoint;
      DWORD	_dwEventCookie;
      ICustomTaskPane& _parent;

      std::list<ICustomTaskPane::ChangeHandler> _visibilityChangeHandlers;
      std::list<ICustomTaskPane::ChangeHandler> _dockStateChangeHandlers;
    };

    class CustomTaskPaneCreator : public ICustomTaskPane
    {
      auto& comAddinImpl()
      {
        return _registrar.server();
      }

      RegisterCom<CustomTaskPaneCtrl> _registrar;
      Office::_CustomTaskPanePtr _pane;
      CComPtr<ComObject<CustomTaskPaneEventHandler>> _paneEvents;


    public:
      CustomTaskPaneCreator(Office::ICTPFactory* ctpFactory, const wchar_t* name)
        : _registrar(
          [](const wchar_t* progId, const GUID& clsid) 
          { 
            return new ComObject<CustomTaskPaneCtrl>(progId, clsid); 
          },
          formatStr(L"%s.CTP", name ? name : L"xlOil").c_str())
      {
        _pane = ctpFactory->CreateCTP(_registrar.progid(), name);
        //_pane = ctpFactory->CreateCTP(L"WinForms.Control.Host.V3", name);
        //_pane = ctpFactory->CreateCTP(L"Paint.Picture", name);
        _paneEvents = new ComObject<CustomTaskPaneEventHandler>(*this);
        _paneEvents->connect(_pane);

        SetWindowPos((HWND)hWnd(), HWND_TOP, 0, 0, 100, 300, SWP_NOMOVE);
      }
      ~CustomTaskPaneCreator()
      {
        _pane->Delete();
      }
      IDispatch* content() const override
      {
        return _pane->ContentControl;
      }
      long hWnd() const override
      {
        //_pane->
        //auto window = Excel::WindowPtr(_pane->Window);
        IOleWindowPtr oleWin(_pane->ContentControl);
        HWND result;
        oleWin->GetWindow(&result);
        return (long)result;
      }
      void setVisible(bool value) override
      { 
        _pane->Visible = value;
      }
      bool getVisible() override
      {
        return _pane->Visible;
      }
      std::pair<int, int> getSize() override
      {
        return std::make_pair(_pane->Width, _pane->Height);
      }
      void setSize(int width, int height) override
      {
        _pane->Width = width;
        _pane->Height = height;
      }
      void addVisibilityChangeHandler(const ChangeHandler& handler) override
      {
        _paneEvents->addVisibilityChangeHandler(handler);
      }
      void addDockStateChangeHandler(const ChangeHandler& handler) override
      {
        _paneEvents->addDockStateChangeHandler(handler);
      }
    };

    ICustomTaskPane* createCustomTaskPane(Office::ICTPFactory* ctpFactory, const wchar_t* name)
    {
      try
      {
        return new CustomTaskPaneCreator(ctpFactory, name);
      }
      XLO_RETHROW_COM_ERROR;
    }
  }
}
