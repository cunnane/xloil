#pragma once
#include <xloil/WindowsSlim.h>

namespace xloil { class Application; }
namespace Excel { struct _Application; }

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Creates a new Excel.Application object and returns a hanging /
    /// detached reference, i.e. no need to AddRef it.
    /// </summary>
    /// <returns></returns>
    Excel::_Application* newApplicationObject();

    /// <summary>
    /// Return the next XLMAIN window after the specified window handle
    /// </summary>
    /// <param name="xlmainHandle"></param>
    /// <returns></returns>
    HWND nextExcelMainWindow(HWND startFrom = 0);

    /// <summary>
    /// Get the Excel.Application corresponding to an XLMAIN window handle.
    /// Returns a hanging / detached reference, i.e. no need to AddRef it.
    /// 
    /// A call to GetActiveObject("Excel.Application") gets the first registered 
    /// instance of Excel which may not be the required one. Hence the need for
    /// this functionality.
    /// </summary>
    Excel::_Application* applicationObjectFromWindow(HWND xlmainHandle);

    /// <summary>
    /// Tries to connect the COM interface if it is not already connected. Returns true 
    /// if the COM interface is available.
    /// </summary>
    bool connectCom();

    void disconnectCom();

    /// <summary>
    /// Performs some trivial COM action to determine if it is likely that 
    /// Excel will accept COM commands (however this is never guaranteed).
    /// </summary>
    bool isComApiAvailable() noexcept;

    Application& attachedApplication();
  }
}