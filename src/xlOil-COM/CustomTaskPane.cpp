
#include "ClassFactory.h"
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/ExcelUI.h>
#include <xloil/Throw.h>
#include <xloil/Log.h>
#include <atlctl.h>
using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  namespace COM
  {

    //public class zCustomTaskPaneCollection
    //{
    //  // Public list of TaskPane items
    //  public List<zCustomTaskPane> Items = new List<zCustomTaskPane>();
    //  private Office.ICTPFactory _paneFactory;
    //};


    // TODO: I don't think we need one of these per pane, just one global one, although we'd have to map back to ICustomTaskPane objects
    class __declspec(novtable) CustomTaskPaneEventHandler
      : public CComObjectRootEx<CComSingleThreadModel>,
      public NoIDispatchImpl<Office::_CustomTaskPaneEvents>
    {
    public:
      CustomTaskPaneEventHandler(ICustomTaskPane& parent, shared_ptr<ICustomTaskPaneHandler> handler)
        : _parent(parent)
        , _handler(handler)
      {}

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
        struct _CustomTaskPane* /*CustomTaskPaneInst*/)
      {
        _handler->onVisible(_parent.getVisible());
        return S_OK;
      }
      HRESULT DockPositionStateChange(
        struct _CustomTaskPane* /*CustomTaskPaneInst*/)
      {
        _handler->onDocked();
        return S_OK;
      }

    private:
      IConnectionPoint* _pIConnectionPoint;
      DWORD	_dwEventCookie;
      ICustomTaskPane& _parent;

      shared_ptr<ICustomTaskPaneHandler> _handler;
    };

    // TODO: do we really need all these interfaces?
    class ATL_NO_VTABLE CustomTaskPaneCtrl :
      public CComObjectRootEx<CComSingleThreadModel>,
      public IDispatchImpl<IDispatch>,
      public CComControl<CustomTaskPaneCtrl>,
      public IOleControlImpl<CustomTaskPaneCtrl>,
      public IOleObjectImpl<CustomTaskPaneCtrl>,
      public IOleInPlaceActiveObjectImpl<CustomTaskPaneCtrl>,
      public IViewObjectExImpl<CustomTaskPaneCtrl>,
      public IOleInPlaceObjectWindowlessImpl<CustomTaskPaneCtrl>
    {
      GUID _clsid;
      std::list<shared_ptr<ICustomTaskPaneHandler>> _handlers;

      unsigned n_bWindowOnly = 1;

    public:
      CustomTaskPaneCtrl(const wchar_t* progId, const GUID& clsid)
        : _clsid(clsid)
      {
      }
      static const CLSID& WINAPI GetObjectCLSID()
      {
        XLO_THROW("Not supported");
      }
      void addHandler(const shared_ptr<ICustomTaskPaneHandler>& events)
      {
        _handlers.push_back(events);
      }
      void destroy()
      {
        for (auto& h : _handlers)
          h->onDestroy();
        _handlers.clear();
      }
      // TODO: should we re-enable windowless mode since we don't need the hwnd
      BEGIN_COM_MAP(CustomTaskPaneCtrl)
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
      END_COM_MAP()

      BEGIN_MSG_MAP(CustomTaskPaneCtrl)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        CHAIN_MSG_MAP(CComControl<CustomTaskPaneCtrl>)
        DEFAULT_REFLECTION_HANDLER()
      END_MSG_MAP()

      // IViewObjectEx
      DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

    public:
      // We need trival implementations of these four methods since we do not have a static CLSID
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
      STDMETHOD(GetUserType)(DWORD dwFormOfType, LPOLESTR* pszUserType)
      {
        return OleRegGetUserType(_clsid, dwFormOfType, pszUserType);
      }
      STDMETHOD(GetMiscStatus)(
        _In_ DWORD dwAspect,
        _Out_ DWORD* pdwStatus)
      {
        return OleRegGetMiscStatus(_clsid, dwAspect, pdwStatus);
      }

      HWND GetActualParent()
      {
        HWND hwndParent = m_hWnd;

        // Get the window associated with the in-place site object,
        // which is connected to this ActiveX control.
        if (m_spInPlaceSite == NULL)
          m_spInPlaceSite->GetWindow(&hwndParent);

        return hwndParent;  
      }


      HRESULT OnSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
      {
        try
        {
          UINT width = LOWORD(lParam);
          UINT height = HIWORD(lParam);
          for (auto& h : _handlers)
            h->onSize(width, height);
        }
        catch (const std::exception& e)
        {
          XLO_ERROR(e.what());
        }
        return S_OK;
      }
    };

    class CustomTaskPaneCreator : public ICustomTaskPane
    {
      Office::_CustomTaskPanePtr _pane;
      std::list<CComPtr<ComObject<CustomTaskPaneEventHandler>>> _paneEvents;
      CComPtr<CustomTaskPaneCtrl> _customCtrl;

    public:
      CustomTaskPaneCreator(
        Office::ICTPFactory& ctpFactory, 
        const wchar_t* name,
        const IDispatch* window,
        const wchar_t* progId)
      {
        // Pasing vtMissing causes the pane to be attached to ActiveWindow
        auto targetWindow = window ? _variant_t(window) : vtMissing;
        if (!progId)
        {
          RegisterCom<CustomTaskPaneCtrl> registrar(
            [](const wchar_t* progId, const GUID& clsid)
            {
              return new ComObject<CustomTaskPaneCtrl>(progId, clsid);
            },
            formatStr(L"%s.CTP", name ? name : L"xlOil").c_str());
          _pane = ctpFactory.CreateCTP(registrar.progid(), name, targetWindow);
          _customCtrl = registrar.server();
        }
        else
          _pane = ctpFactory.CreateCTP(progId, name, targetWindow);
      }
      ~CustomTaskPaneCreator()
      {
        destroy();
      }
      IDispatch* content() const override
      {
        return _pane->ContentControl;
      }
      ExcelWindow window() const override
      {
        return ExcelWindow(Excel::WindowPtr(_pane->Window));
      }

      size_t parentWindowHandle() const override
      {
        HWND parent = 0;
        if (_customCtrl)
          parent = _customCtrl->GetActualParent();
        else
        {
          IOleWindowPtr oleWin(_pane->ContentControl);
          HWND result;
          oleWin->GetWindow(&parent);
        }
        constexpr wchar_t target[] = L"NUIPane";      // could be MsoWorkPane
        constexpr auto len = 1 + _countof(target);
        wchar_t winClass[len + 1];
        //  Ensure that class_name is always null terminated for safety.
        winClass[len] = 0;

        do 
        {
          auto hwnd = parent;
          parent = ::GetParent(hwnd);
          if (parent == hwnd)
            XLO_THROW(L"Failed to find parent window with class {}", target);
          ::GetClassName(parent, winClass, len);
        } while (wcscmp(target, winClass) != 0);

        return (size_t)parent;
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
      DockPosition getPosition() const override
      {
        return DockPosition(_pane->DockPosition);
      }
      void setPosition(DockPosition pos) override
      {
        _pane->DockPosition = (Office::MsoCTPDockPosition)pos;
      }

      std::wstring getTitle() const
      {
        return _pane->Title.GetBSTR();
      }
      void destroy() const
      {
        if (_customCtrl)
          _customCtrl->destroy();
        _pane->Delete();
      }
      void addEventHandler(const std::shared_ptr<ICustomTaskPaneHandler>& events) override
      {
        _paneEvents.push_back(new ComObject<CustomTaskPaneEventHandler>(*this, events));
        _paneEvents.back()->connect(_pane);
        if (_customCtrl)
          _customCtrl->addHandler(events);
      }
      
    };

    ICustomTaskPane* createCustomTaskPane(
      Office::ICTPFactory& ctpFactory, 
      const wchar_t* name,
      const IDispatch* window,
      const wchar_t* progId)
    {
      try
      {
        return new CustomTaskPaneCreator(ctpFactory, name, window, progId);
      }
      XLO_RETHROW_COM_ERROR;
    }
  }
}
