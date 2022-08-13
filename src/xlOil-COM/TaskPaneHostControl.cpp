#include "TaskPaneHostControl.h"
#include "ClassFactory.h"
#include <xloil/Log.h>
#include <atlbase.h>
#include <atlctl.h>


namespace xloil
{
  namespace COM
  {
    // TODO: do we really need all these interfaces?
    class ATL_NO_VTABLE TaskPaneHostControl :
      public CComObjectRootEx<CComSingleThreadModel>,
      public IDispatchImpl<IDispatch>,
      public CComControl<TaskPaneHostControl>,
      public IOleControlImpl<TaskPaneHostControl>,
      public IOleObjectImpl<TaskPaneHostControl>,
      public IOleInPlaceActiveObjectImpl<TaskPaneHostControl>,
      public IViewObjectExImpl<TaskPaneHostControl>,
      public IOleInPlaceObjectWindowlessImpl<TaskPaneHostControl>,
      public ITaskPaneHostControl
    {
      HWND _attachedWindow = 0;
      HWND _previousParent = 0;
      LONG_PTR _previousWindowStyle = 0;

      static const GUID _clsid;

    public:
      TaskPaneHostControl() noexcept
      {
        m_bWindowOnly = 1;
      }

      ~TaskPaneHostControl() noexcept
      {
        detach();
      }

      void detach() noexcept
      {
        if (_attachedWindow == 0)
          return;
        auto parent = ::SetParent(_attachedWindow, _previousParent);
        ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, _previousWindowStyle);
        RemoveWindowSubclass(parent, SubclassWndProc, (UINT_PTR)this);
        _attachedWindow = 0;
      }

      static const CLSID& WINAPI GetObjectCLSID()
      {
        return _clsid;
      }

      /// <summary>
      /// Given a window handle, make it a frameless child of the one of
      /// our parent windows
      /// </summary>
      void AttachWindow(HWND hwnd) override
      {
        _attachedWindow = hwnd;
        auto parent = GetAttachableParent();
        _previousParent = ::SetParent(_attachedWindow, parent);

        // Subclass the parent so that when we get a WM_DESTROY, we can unattach
        // the window. Some GUI toolkits (e.g. Qt) object to their windows being
        // destroyed from unexpected sources.
        SetWindowSubclass(parent, SubclassWndProc, (UINT_PTR)this, (DWORD_PTR)nullptr);

        XLO_DEBUG("Custom task pane host control attached to window {0x}. Previous parent {0x}",
          (size_t)hwnd, (size_t)_previousParent);

        _previousWindowStyle = ::GetWindowLongPtr(_attachedWindow, GWL_STYLE);
        auto style = (_previousWindowStyle | WS_CHILD) & ~WS_THICKFRAME & ~WS_CAPTION;
        ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, style);

        // Set the z-order and reposition the child at the top left of the parent
        // Not sure why we need both of these!
        ::SetWindowPos(m_hWnd, _attachedWindow, 0, 0, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
        ::SetWindowPos(_attachedWindow, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
      }

      BEGIN_COM_MAP(TaskPaneHostControl)
        COM_INTERFACE_ENTRY(IDispatch)
        COM_INTERFACE_ENTRY(IViewObjectEx)
        COM_INTERFACE_ENTRY(IViewObject2)
        COM_INTERFACE_ENTRY(IViewObject)
        COM_INTERFACE_ENTRY(IOleInPlaceObject)
        COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObject)
        COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
        COM_INTERFACE_ENTRY(IOleControl)
        COM_INTERFACE_ENTRY(IOleObject)
        COM_INTERFACE_ENTRY(ITaskPaneHostControl)
      END_COM_MAP()

      BEGIN_MSG_MAP(TaskPaneHostControl)
        MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnPosChanging)
        CHAIN_MSG_MAP(CComControl<TaskPaneHostControl>)
        DEFAULT_REFLECTION_HANDLER()
      END_MSG_MAP()

      // IViewObjectEx
      DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

      DECLARE_WND_CLASS(_T("xlOilAXHostControl"))

    private:
      HWND GetAttachableParent()
      {
        // We can't link GUI window to just any old parent. In particular not
        // m_hWnd of this class. This is because the DPI awareness for this
        // window is set to System whereas the GUI toolkit root windows
        // are at Per-Monitor or better; this causes GetParent to fail. Because
        // this window's parent is also System awareness, we can't make our 
        // window Per-Monitor aware. Even if we call SetThreadDpiHostingBehavior
        // with DPI_HOSTING_BEHAVIOR_MIXED - it just doesn't work. Instead we 
        // walk up the parent chain unti we find "NUIPane" which experimentation
        // has shown will be a suitable parent.
        // 
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

      static LRESULT CALLBACK SubclassWndProc(HWND hWnd, UINT uMsg,
        WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass,
        DWORD_PTR /*dwRefData*/)
      {
        switch (uMsg)
        {
        case WM_DESTROY:
          ((TaskPaneHostControl*)uIdSubclass)->detach();
        }
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
      }
    };

    // {EBE296D2-2373-437C-9FF5-934865BAB572}
    const GUID TaskPaneHostControl::_clsid =
    { 0xebe296d2, 0x2373, 0x437c, { 0x9f, 0xf5, 0x93, 0x48, 0x65, 0xba, 0xb5, 0x72 } };

    const wchar_t* taskPaneHostControlProgId()
    {
      static RegisterCom registrar(
        []() { return (IDispatch*)new CComObject<TaskPaneHostControl>(); },
        L"xlOilAXHostControl", 
        &TaskPaneHostControl::GetObjectCLSID());
      return registrar.progid();
    }
  }
}