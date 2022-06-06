#include "Connect.h"
#include "ComEventSink.h"
#include <xloil/Events.h>
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/AppObjects.h>
#include <oleacc.h> // AccessibleObjectFromWindow
#include <memory>

using std::make_shared;
using std::wstring;

namespace xloil
{
  namespace COM
  {
    HWND nextExcelMainWindow(HWND xlmainHandle)
    {
      return FindWindowExA(0, xlmainHandle, "XLMAIN", NULL);
    }

    Excel::_Application* newApplicationObject()
    {
      CoInitialize(NULL); // It's safe to call repeatedly
      Excel::_ApplicationPtr app;
      auto hr = app.CreateInstance(
        __uuidof(Excel::Application),
        NULL,
        CLSCTX_LOCAL_SERVER);
      return hr == S_OK ? app.Detach() : nullptr;
    }

    Excel::_Application* applicationObjectFromWindow(HWND xlmainHandle)
    {
      CoInitialize(NULL);
      // Based on discussion here:
      // https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
      auto hwnd2 = FindWindowExA(xlmainHandle, 0, "XLDESK", NULL);
      auto hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", NULL);
      Excel::Window* pWindow = NULL;
      if (AccessibleObjectFromWindow(hwnd3, (DWORD)OBJID_NATIVEOM,
        __uuidof(IDispatch),
        (void**)&pWindow) == S_OK)
        return pWindow->Application.Detach();
      return nullptr;
    }

    namespace
    {
      class COMConnector
      {
      public:
        COMConnector()
        {
          try
          {

            auto windowHandle = callExcel(msxll::xlGetHwnd);
            // This conversion to 32-bit is OK even in x64 because the 
            // window handle is an index into an array, not a pointer. 
#pragma warning(disable: 4312)
            _excelWindowHandle = (HWND)windowHandle.get<int>();

            _xlApp = applicationObjectFromWindow(_excelWindowHandle);
            _handler = COM::createEventSink(_xlApp);
            
            XLO_DEBUG(L"Made COM connection to Excel at '{}' with hwnd={}",
              (const wchar_t*)_xlApp->Path, (size_t)_excelWindowHandle);
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

    Excel::_Application& attachedApplication()
    {
      if (!theComConnector)
        throw ComConnectException("COM Connection not ready");
      return theComConnector->excelApp();
    }
  }
}