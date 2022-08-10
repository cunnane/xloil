#include "ClassFactory.h"
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/AppObjects.h>
#include <xlOil/ExcelUI.h>
#include <xloil/Throw.h>
#include <xloil/Log.h>
#include <xloil/State.h>
#include <atlbase.h>
#include <atlctl.h>

using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  namespace COM
  {
    class CustomTaskPaneEventHandler
      : public ComEventHandler<
          NoIDispatchImpl<ComObject<Office::_CustomTaskPaneEvents>>, Office::_CustomTaskPaneEvents>
    {
    public:
      CustomTaskPaneEventHandler(
        ICustomTaskPane& parent, 
        shared_ptr<ICustomTaskPaneEvents> handler)
        : _parent(parent)
        , _handler(handler)
      {}

      virtual ~CustomTaskPaneEventHandler() noexcept
      {}

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

      void destroy() noexcept
      {
        try
        {
          _handler->onDestroy();
          disconnect();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR(e.what());
        }
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
      ICustomTaskPane& _parent;
      shared_ptr<ICustomTaskPaneEvents> _handler;
    };
    
    // TODO: do we really need all these interfaces?
    class ATL_NO_VTABLE WindowHostControl :
      public CComObjectRootEx<CComSingleThreadModel>,
      public IDispatchImpl<IDispatch>,
      public CComControl<WindowHostControl>,
      public IOleControlImpl<WindowHostControl>,
      public IOleObjectImpl<WindowHostControl>,
      public IOleInPlaceActiveObjectImpl<WindowHostControl>,
      public IViewObjectExImpl<WindowHostControl>,
      public IOleInPlaceObjectWindowlessImpl<WindowHostControl>
    {
      GUID _clsid;
      HWND _attachedWindow = 0;
      LONG_PTR _previousWindowStyle = 0;

    public:
      WindowHostControl() noexcept
      { 
        m_bWindowOnly = 1;
      }
      ~WindowHostControl() noexcept
      {
        ::SetParent(_attachedWindow, NULL);
        ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, _previousWindowStyle);
      }
      void init(const GUID& clsid) noexcept
      {
        _clsid = clsid;
      }
      static const CLSID& WINAPI GetObjectCLSID()
      {
        XLO_THROW("Not supported");
      }

      /// <summary>
      /// Given a window handle, make it a frameless child of the one of
      /// our parent windows
      /// </summary>
      void attachWindow(size_t hwnd)
      {
        _attachedWindow = (HWND)hwnd;
        ::SetParent(_attachedWindow, GetAttachableParent());

        _previousWindowStyle = ::GetWindowLongPtr(_attachedWindow, GWL_STYLE);
        auto style = (_previousWindowStyle | WS_CHILD) & ~WS_THICKFRAME & ~WS_CAPTION;
        ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, style);

        // Set the z-order and reposition the child at the top left of the parent
        // TODO: Not sure we need both of these!
        ::SetWindowPos(m_hWnd, _attachedWindow, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        ::SetWindowPos(_attachedWindow, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
      }
      DECLARE_WND_CLASS(_T("xlOilAXControl"))

      BEGIN_COM_MAP(WindowHostControl)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IViewObjectEx)
        COM_INTERFACE_ENTRY(IViewObject2)
        COM_INTERFACE_ENTRY(IViewObject)
        COM_INTERFACE_ENTRY(IOleInPlaceObject)
        COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObject)
        COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
        COM_INTERFACE_ENTRY(IOleControl)
        COM_INTERFACE_ENTRY(IOleObject)
      END_COM_MAP()

      BEGIN_MSG_MAP(WindowHostControl)
        MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnPosChanging)
        CHAIN_MSG_MAP(CComControl<WindowHostControl>)
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

      // We can't link GUI window to just any old parent. In particular not
      // m_hWnd of this class. This is because the DPI awareness for this
      // window is set to System whereas the GUI toolkit root windows
      // are at Per-Monitor or better; this causes GetParent to fail. Because
      // this window's parent is also System awareness, we can't make our 
      // window Per-Monitor aware. Even if we call SetThreadDpiHostingBehavior
      // with DPI_HOSTING_BEHAVIOR_MIXED - it just doesn't work. Instead we 
      // walk up the parent chain unti we find "NUIPane" which experimentation
      // has shown will be a suitable parent.
      HWND GetAttachableParent()
      {
        // We will walk up the window stack until we find the target below.
        // TODO: could be MsoWorkPane or NetUIHWND?
        constexpr wchar_t target[] = L"NUIPane";  

        constexpr auto len = 1 + _countof(target);

        // Create a buffer for window class names; doesn't need to be larger
        // than the target name.  Ensure the buffer is always null terminated 
        // for safety.
        wchar_t winClass[len + 1];
        winClass[len] = 0;

        HWND parent = m_hWnd;
        do
        {
          auto hwnd = parent;
          parent = ::GetParent(hwnd);
          if (parent == hwnd || parent == NULL)
            XLO_THROW(L"Failed to find parent window with class {}", target);
          ::GetClassName(parent, winClass, len);
        } while (wcscmp(target, winClass) != 0);

        return parent;
      }

      HRESULT OnPosChanging(UINT message, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
      {
        auto windowPos = (WINDOWPOS*)lParam;
        if (_attachedWindow != 0 && windowPos->cx > 0 && windowPos->cy > 0)
        {
          ::SetWindowPos(
            _attachedWindow, 
            0, 0, 0, 
            windowPos->cx, windowPos->cy, 
            SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        bHandled = true;
        return DefWindowProc(message, wParam, lParam);
      }

    };

    class CustomTaskPaneCreator : public ICustomTaskPane
    {
      Office::_CustomTaskPanePtr _pane;
      CComPtr<CustomTaskPaneEventHandler> _paneEvents;
      CComPtr<WindowHostControl> _hostingControl;

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
          RegisterCom<WindowHostControl> registrar(
            [](const wchar_t* progId, const GUID& clsid)
            {
              auto p = new CComObject<WindowHostControl>;
              p->init(clsid);
              return p;
            },
            formatStr(L"%s.CTP", name ? name : L"xlOil").c_str());
          _pane = ctpFactory.CreateCTP(registrar.progid(), name, targetWindow);
          _hostingControl = registrar.server();
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

      void destroy()
      {
        if (_paneEvents)
        {
          _paneEvents->destroy();
          _paneEvents.Release();
        }
        if (_hostingControl)
          _hostingControl.Release();
        _pane->Delete();
      }

      void listen(const std::shared_ptr<ICustomTaskPaneEvents>& events) override
      {
        _paneEvents = CComPtr<CustomTaskPaneEventHandler>(
          new CustomTaskPaneEventHandler(*this, events));
        _paneEvents->connect(_pane);
      }

      void attach(size_t hwnd) override
      {
        if (_hostingControl)
          _hostingControl->attachWindow(hwnd);
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
