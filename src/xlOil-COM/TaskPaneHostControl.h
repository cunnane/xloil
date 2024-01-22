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
      /// Reparents the specified window into the host control, making it
      /// a fixed position borderless child.
      /// </summary>
      virtual void AttachWindow(HWND hwnd) = 0;
      /// <summary>
      /// Specifies a handler for the onDestroy event called when the pane
      /// receives a WM_DESTROY message. Pass a null pointer to detatch.
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