#include <comdef.h>

namespace xloil
{
  namespace COM
  {
    struct __declspec(uuid("2ADAD4E5-0793-4151-8D29-07B05C4B0557"))
      ITaskPaneHostControl : public IUnknown
    {
      virtual void AttachWindow(HWND hwnd) = 0;
    };

    /// <summary>
    /// Registers the task pane hosting COM control and returns the progid
    /// </summary>
    /// <returns></returns>
    const wchar_t* taskPaneHostControlProgId();
  }
}