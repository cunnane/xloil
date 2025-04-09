#define WIN32_LEAN_AND_MEAN
#include <atlbase.h>
#include <atlctl.h>

#include "TaskPaneHostControl.h"
#include "ClassFactory.h"
#include <xloil/Log.h>
#include <xloil/ExcelUI.h>
#include <xloil/State.h>

#pragma comment(lib, "comctl32.lib")

using std::shared_ptr;

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Implements much hackery, largely discovered by experimentation to ensure
    /// that a target window is displayed over a task pane.  The "shadowed" window
    /// attachment style in particular uses some steps that may be superfluous but
    /// appear to make things work
    /// </summary>
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
      HWND _workPane = 0;
      HWND _parent = 0;
      LONG_PTR _previousWindowStyle = 0;
      shared_ptr<ICustomTaskPaneEvents> _destroyHandler = nullptr;
      HWND _excelHwnd = 0;
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

        XLO_DEBUG("Detaching and destroying task pane with hwnd={:x}", (size_t)_attachedWindow);

        ::RemoveWindowSubclass(_parent,
          parentSubclassProc, (UINT_PTR)this);

        if (attachedAsParent())
        {
          // Restore parent
          ::SetParent(_attachedWindow, _previousParent);
          
        }
        else
        {
          ::RemoveWindowSubclass(_excelHwnd,
            handleMoveSubclassProc, (UINT_PTR)this);
        }

        // Restore previous window style
        ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, _previousWindowStyle);

        // Hide the window so it doesn't jump around the screen before its creator can destroy it
        ::ShowWindow(_attachedWindow, SW_HIDE);
 
        _attachedWindow = 0;

        if (_destroyHandler)
        {
          try
          {
            _destroyHandler->onDestroy();
          }
          catch (const std::exception& e)
          {
            XLO_ERROR(e.what());
          }
        }
      }

      static const CLSID& WINAPI GetObjectCLSID()
      {
        return _clsid;
      }

      void AttachWindow(HWND hwnd, bool asParent) override
      {
        XLO_DEBUG("Task pane host attaching to {:x}", (size_t)hwnd);

        _attachedWindow = hwnd; 
        _previousWindowStyle = ::GetWindowLongPtr(_attachedWindow, GWL_STYLE);
        _parent = GetAttachableParent();

        // Subclass the parent so that Note the subclass id is 'this' which
        // we cast back in the WndProc.
        SetWindowSubclass(_parent,
          parentSubclassProc, (UINT_PTR)this, (DWORD_PTR)nullptr);

        if (asParent)
        {
          _previousParent = ::SetParent(_attachedWindow, _parent);

          XLO_DEBUG("Task pane host control adopted window {:x}. Previous parent {:x}",
            (size_t)hwnd, (size_t)_previousParent);

          // Change style to frameless child
          auto style = (_previousWindowStyle | WS_CHILD) & ~WS_THICKFRAME & ~WS_CAPTION;
          ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, style);

          // Set the z-order and reposition the child at the top left of the parent.
          ::SetWindowPos(m_hWnd, _attachedWindow, 0, 0, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
          ::SetWindowPos(_attachedWindow, 0, 0, 0, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOZORDER);
        }
        else
        {
          auto style = _previousWindowStyle & ~WS_THICKFRAME & ~WS_CAPTION &
            ~WS_EX_APPWINDOW & ~WS_EX_TOOLWINDOW;
          ::SetWindowLongPtr(_attachedWindow, GWL_STYLE, style);

          // Give the window an owner to stop it appearing in the taskbar
          ::SetWindowLongPtr(_attachedWindow, GWLP_HWNDPARENT, (LONG_PTR)_parent);

          // Show/hide the window to update teh style changes
          ::ShowWindow(_attachedWindow, SW_HIDE);
          ::ShowWindow(_attachedWindow, SW_SHOW);

          moveAttachedWindow();

          _excelHwnd = GetExcelHWND(_parent);

          ::SetWindowSubclass(_excelHwnd, 
            handleMoveSubclassProc, (UINT_PTR)this, (DWORD_PTR)nullptr);
        }
      }

      void AttachDestroyHandler(const std::shared_ptr<ICustomTaskPaneEvents>& handler) override
      {
        _destroyHandler = handler;
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
      bool attachedAsParent() const 
      {
        return _previousParent != nullptr;
      }

      /// <summary>
      /// Returns the XLMAIN hwnd associated with the given window - Excel
      /// can have multiple XLMAIN top level windows open, so this function
      /// walks up the stack to find the target
      /// </summary>
      /// <param name="from"></param>
      /// <returns></returns>
      HWND GetExcelHWND(HWND from)
      {
        constexpr wchar_t target[] = L"XLMAIN";
        return getWindowByClass(from, target, 1 + _countof(target));
      }

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
        constexpr wchar_t target[] = L"NUIPane";
        return getWindowByClass(m_hWnd, target, 1 + _countof(target));
      }

      HWND getWindowByClass(HWND from, const wchar_t* className, size_t len = 0)
      {
        if (len == 0)
          len = wcslen(className);

        assert(len < 256);

        // Create a buffer for window class names. Ensure the buffer is always null 
        // terminated for safety.
        wchar_t winClass[256];
        winClass[255] = 0;

        HWND parent = from;
        do
        {
          auto hwnd = parent;
          parent = ::GetParent(hwnd);
          if (parent == hwnd || parent == NULL)
            XLO_THROW(L"Failed to find parent window with class {}", className);
          ::GetClassName(parent, winClass, (int)len);
        } while (wcscmp(className, winClass) != 0);

        return parent;
      }
      
      /// <summary>
      /// Moves the attached window to occupy the same space as the parent panel
      /// and be on top of the Excel main window.
      /// </summary>
      void moveAttachedWindow()
      {
        RECT attached;
        ::GetWindowRect(_parent, &attached);

        ::SetWindowPos(_attachedWindow,
          0,
          attached.left, attached.top,
          0, 0,
          SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER | SWP_FRAMECHANGED);

        ::SetWindowPos(_attachedWindow,
          HWND_TOP,
          0, 0, 0, 0,
          SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOMOVE);
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
          if (!attachedAsParent())
            moveAttachedWindow();
        }
        bHandled = true;
        return (HRESULT)DefWindowProc(message, wParam, lParam);
      }

      static LRESULT CALLBACK parentSubclassProc(
        HWND hWnd, UINT uMsg,
        WPARAM wParam, LPARAM lParam, 
        UINT_PTR uIdSubclass,
        DWORD_PTR /*dwRefData*/)
      {
        auto self = ((TaskPaneHostControl*)uIdSubclass);
        switch (uMsg)
        {
        //case 0xD00:
        case WM_CHILDACTIVATE:
        {
          if (!self->attachedAsParent())
          {
            // When a taskpane is popped out a new top level window is created to parent 
            // it, this window is destroyed when the task pane is docked so we need to keep
            // up with the latest parent
            auto workpane = ::GetParent(::GetParent(self->_parent));
            if (workpane != self->_workPane)
            {
              ::RemoveWindowSubclass(self->_workPane,
                handlePositionSubclassProc, (UINT_PTR)self);
              self->_workPane = workpane;
              ::SetWindowSubclass(self->_workPane,
                handlePositionSubclassProc, (UINT_PTR)self, (DWORD_PTR)nullptr);
              self->moveAttachedWindow();
            }
          }
          break;
        }
        case WM_DESTROY:
          // When we get a WM_DESTROY we dettach the window. Some GUI toolkits (e.g. Qt)
          // object to their windows being destroyed from unexpected sources and dump a
          // core to register their displeasure.
          ((TaskPaneHostControl*)uIdSubclass)->detach();
          break;
        case WM_WINDOWPOSCHANGING:
        {
          // We get here when the Office backstage view is opened, which hides the
          // task panes
          if (!self->attachedAsParent())
          {
            WINDOWPOS* p = (WINDOWPOS*)lParam;
            if ((p->flags & SWP_HIDEWINDOW) != 0)
              ::ShowWindow(self->_attachedWindow, SW_HIDE);

            // You might think that SWP_SHOWWINDOW should be the correct flag
            // here, but experimentation shows that SWP_NOCOPYBITS works
            else if ((p->flags & SWP_NOCOPYBITS) != 0)
              ::ShowWindow(self->_attachedWindow, SW_SHOW);
          }
          break;
        }
        }
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
      }

      static LRESULT CALLBACK handleMoveSubclassProc(
        HWND hWnd, UINT uMsg,
        WPARAM wParam, LPARAM lParam, 
        UINT_PTR uIdSubclass,
        DWORD_PTR /*dwRefData*/)
      {
        auto self = ((TaskPaneHostControl*)uIdSubclass);
        auto retVal = DefSubclassProc(hWnd, uMsg, wParam, lParam);

        // If anything interesting happens, i.e. the window moves, we create
        // a callback to moveAttachedWindow by which time the window will be in
        // its new position as expected by moveAttachedWindow.
        switch (uMsg)
        {
        case WM_USER + 643:
          self->moveAttachedWindow();
          break;
        case WM_SIZE:
        case WM_MOVE:
          ::PostMessage(hWnd, WM_USER + 643, 0, 0);
          break;
        }
        return retVal;
      }

      static LRESULT CALLBACK handlePositionSubclassProc(
        HWND hWnd, UINT uMsg,
        WPARAM wParam, LPARAM lParam,
        UINT_PTR uIdSubclass,
        DWORD_PTR /*dwRefData*/)
      {
        auto self = ((TaskPaneHostControl*)uIdSubclass);
        auto retVal = DefSubclassProc(hWnd, uMsg, wParam, lParam);

        switch (uMsg)
        {
        case WM_USER + 643:
          self->moveAttachedWindow();
          break;
        case WM_SIZE:
        case WM_MOVE:
        {
          ::PostMessage(hWnd, WM_USER + 643, 0, 0);
          break;
        }
        case WM_WINDOWPOSCHANGING:
        {
          WINDOWPOS* p = (WINDOWPOS*)lParam;
          if ((p->flags & SWP_NOZORDER) == 0 && ::IsWindowVisible(self->_attachedWindow))
          {
            // Fix the z order
            ::SetWindowPos(self->_attachedWindow,
              p->hwndInsertAfter,
              0, 0, 0, 0,
              SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
            p->hwndInsertAfter = self->_attachedWindow;
          }
          break;
        }
        }
        return retVal;
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