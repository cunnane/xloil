#include "Connect.h"
#include "ComEventSink.h"
#include <xloil/Events.h>
#include <xloil/ExcelCall.h>

#include "ExcelTypeLib.h"

#include <set>
#include <memory>

using std::make_shared;
using std::wstring;
using std::set;


Excel::_ApplicationPtr getExcelObjFromWindow(HWND xlmainHandle)
{
  // Based on discussion here:
  // https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
  HWND hwnd = nullptr, hwnd2, hwnd3;
  hwnd2 = FindWindowExA(xlmainHandle, 0, "XLDESK", NULL);
  hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", NULL);
  Excel::Window* pWindow = NULL;
  if (AccessibleObjectFromWindow(hwnd3, OBJID_NATIVEOM, __uuidof(IDispatch), (void**)&pWindow) == S_OK)
    return pWindow->Application;
  return nullptr;
}

/// <summary>
/// A naive GetActiveObject("Excel.Application") gets the first registered 
/// instance of Excel which may not be our instance. Instead we get the one
/// corresponding to the window handle we get from xlGetHwnd.
/// </summary>
/// 
Excel::_ApplicationPtr getExcelInstance(HWND xlmainHandle)
{
  auto hwndCurrent = ::GetForegroundWindow();

  // We switch focus away from Excel because that increases
  // the chances of the instance adding itself to the running
  // object table. It isn't determinimistic though so we have
  // to give it a few tries.
  // This apparently bizarre approach is suggested here
  // https://support.microsoft.com/en-za/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application
  for (auto moreTries = 0; moreTries < 15; ++moreTries)
  {
    ::SetForegroundWindow(hwndCurrent);
    auto ptr = getExcelObjFromWindow(xlmainHandle);
    if (ptr)
      return ptr;

    // Chances of an explorer window being available are good
    auto explorerWindow = FindWindow(L"CabinetWClass", nullptr);
    ::SetForegroundWindow(explorerWindow);
    Sleep(300);
  }

  // Need to ensure the foreground window is restored
  ::SetForegroundWindow(hwndCurrent);
  XLO_THROW("Failed to get Excel COM object");
}

namespace xloil
{
  namespace
  {
    class COMConnector
    {
    public:
      COMConnector()
      {
        try
        {
          CoInitialize(NULL);
          auto windowHandle = callExcel(msxll::xlGetHwnd);
          // This conversion to 32-bit is OK even in x64 because the 
          // window handle is an index into an array, not a pointer. 
#pragma warning(disable: 4312)
          _excelWindowHandle = (HWND)windowHandle.toInt();

          Excel::_Application* p = ExcelApp();
          _handler = COM::createEventSink(p);
        }
        catch (_com_error& error)
        {
          XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
        }
      }

      ~COMConnector()
      {
        _handler.reset();
        CoUninitialize();
      }

      const Excel::_ApplicationPtr& ExcelApp() 
      { 
        if (!_xlApp)
          _xlApp = getExcelInstance(_excelWindowHandle);
        return _xlApp;
      }

    private:
      Excel::_ApplicationPtr _xlApp;
      std::shared_ptr<Excel::AppEvents> _handler;
      HWND _excelWindowHandle;
    };

    struct RegisterMe
    {
      RegisterMe() {}
      
      COMConnector* connect()
      {
        if (!connector)
        {
          connector = new COMConnector();
          _handler = Event::AutoClose() += [this]() { this->disconnect(); };
        }
        return connector;
      }

      void disconnect()
      {
        if (connector)
        {
          Event::AutoClose() -= _handler;
          delete connector;
          connector = nullptr;
        }
      }

      typename std::remove_reference<decltype(Event::AutoClose())>::type::handler_id _handler;

    private:
      COMConnector* connector;
    } theInstance;
  }

  void reconnectCOM()
  {
    theInstance.disconnect();
    theInstance.connect();
  }

  Excel::_Application& excelApp()
  {
    return *theInstance.connect()->ExcelApp();
  }

  bool checkWorkbookIsOpen(const wchar_t* workbookName)
  {
    // See other possibility here. Seems a bit crazy?
    // https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
    try
    {
      auto workbook = excelApp().Workbooks->GetItem(_variant_t(workbookName));
      return !!workbook;
    }
    catch (_com_error& error)
    {
      if (error.Error() == DISP_E_BADINDEX)
        return false;
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }
}