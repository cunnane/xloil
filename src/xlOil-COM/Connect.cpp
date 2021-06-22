#include "Connect.h"
#include "ComEventSink.h"
#include <xloil/Events.h>
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelTypeLib.h>
#include <oleacc.h> // AccessibleObjectFromWindow
#include <memory>

using std::make_shared;
using std::wstring;


Excel::_ApplicationPtr getExcelObjFromWindow(HWND xlmainHandle)
{
  // Based on discussion here:
  // https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
  auto hwnd2 = FindWindowExA(xlmainHandle, 0, "XLDESK", NULL);
  auto hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", NULL);
  Excel::Window* pWindow = NULL;
  if (AccessibleObjectFromWindow(hwnd3, (DWORD)OBJID_NATIVEOM, 
                                 __uuidof(IDispatch), 
                                 (void**)&pWindow) == S_OK)
    return pWindow->Application;
  return nullptr;
}


namespace xloil {
  namespace COM
  {
    namespace
    {
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
        // object table. It isn't determinimistic though so the caller
        // may have to retry if it fails.
        // This apparently bizarre approach is suggested here
        // https://support.microsoft.com/en-za/help/238610/getobject-or-getactiveobject-cannot-find-a-running-office-application

        ::SetForegroundWindow(hwndCurrent);
        auto ptr = getExcelObjFromWindow(xlmainHandle);
        if (ptr)
          return ptr;

        // Chances of an explorer window being available are good
        auto explorerWindow = FindWindow(L"CabinetWClass", nullptr);
        ::SetForegroundWindow(explorerWindow);

        // Need to ensure the foreground window is restored
        ::SetForegroundWindow(hwndCurrent);
        throw ComConnectException("Failed to get Excel COM object");
      }

      class COMConnector
      {
      public:
        COMConnector()
        {
          try
          {
            CoInitialize(NULL); // It's safe to call again if we're retrying here
            auto windowHandle = callExcel(msxll::xlGetHwnd);
            // This conversion to 32-bit is OK even in x64 because the 
            // window handle is an index into an array, not a pointer. 
#pragma warning(disable: 4312)
            _excelWindowHandle = (HWND)windowHandle.toInt();

            _xlApp = getExcelInstance(_excelWindowHandle);
            _handler = COM::createEventSink(_xlApp);
          }
          catch (_com_error& error)
          {
            throw ComConnectException(
              utf16ToUtf8(
                fmt::format(L"COM Error {0:#x}: {1}",
                (size_t)error.Error(), error.ErrorMessage())).c_str());
          }
        }

        ~COMConnector()
        {
          _handler.reset();
          CoUninitialize();
        }

        const Excel::_ApplicationPtr& excelApp() const { return _xlApp; }
       
      private:
        Excel::_ApplicationPtr _xlApp;
        std::shared_ptr<Excel::AppEvents> _handler;
        HWND _excelWindowHandle;
      };

      std::unique_ptr<COMConnector> theComConnector;
    }

    void connectCom()
    {
      theComConnector.reset(new COMConnector());
    }

    void disconnectCom()
    {
      theComConnector.reset();
    }

    bool isComApiAvailable() noexcept
    {
      if (!theComConnector)
        return false;
      
      // Do some random COM thing - is this the fastest thing?
      VARIANT_BOOL temp;
      auto result = theComConnector->excelApp()->get_EnableEvents(&temp);
      return (SUCCEEDED(result));
    }

    Excel::_Application& excelApp()
    {
      if (!theComConnector)
        throw ComConnectException("COM Connection not ready");
      return theComConnector->excelApp();
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
}