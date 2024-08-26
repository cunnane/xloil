#include <comdef.h>
#include <memory>

namespace xloil
{
  class ICustomTaskPaneEvents;

  namespace COM
  {
    struct __declspec(uuid("2ADAD4E5-0793-4151-8D29-07B05C4B0557"))
      ITaskPaneHostControl : public IUnknown
    {
      /// <summary>
      /// Makes the specified borderless and keeps it in the right place over
      /// the host control.
      ///  
      /// There are two attachment styles:
      ///   "as parent": much simpler, the target window gets a task pane parent,
      ///   so moving, z-order and visibility are handled automatically by Win32
      /// 
      ///   "shadowed": xlOil keeps the window with the correct location, z-order 
      ///   and visibility 
      /// 
      /// The "shadowed" style exists because certain window managers (Wx) misbehave
      /// if their windows are re-parented.
      /// </summary>
      virtual void AttachWindow(HWND hwnd, bool asParent=true) = 0;
      /// <summary>
      /// Specifies a handler for the onDestroy event called when the pane
      /// receives a WM_DESTROY message. Pass a null pointer to detatch the
      /// handler.
      /// </summary>
      virtual void AttachDestroyHandler(
        const std::shared_ptr<ICustomTaskPaneEvents>& handler) = 0;
    };

    /// <summary>
    /// Registers the task pane hosting COM control and returns the progid
    /// </summary>
    /// <returns></returns>
    const wchar_t* taskPaneHostControlProgId();
  }
}