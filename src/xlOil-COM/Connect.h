#pragma once
#include <xloil/Throw.h>
#include <xloil/WindowsSlim.h>

namespace Excel {
  struct _Application;
}

namespace xloil
{
  namespace COM
  {
    HWND nextExcelMainWindow(HWND xlmainHandle = 0);

    /// <summary>
    /// A naive GetActiveObject("Excel.Application") gets the first registered 
    /// instance of Excel which may not be our instance. Instead we get the one
    /// corresponding to the window handle we get from xlGetHwnd.
    /// </summary>
    /// 
    Excel::_Application* applicationObjectFromWindow(HWND xlmainHandle);

    void connectCom();
    void disconnectCom();

    bool isComApiAvailable() noexcept;

    Excel::_Application& attachedExcelApp();
  }
}