#include "Connect.h"
#include "ComEventSink.h"
#include <xloil/AppObjects.h>
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
      {
        auto result = pWindow->Application.Detach();
        pWindow->Release();
        return result;
      }
      return nullptr;
    }

    namespace
    {
      class COMConnector
      {
      public:
        COMConnector()
          : _excelWindowHandle((HWND)App::internals().hWnd)
          , _xlApp((size_t)_excelWindowHandle)
        {
          try
          {
            _handler = COM::createEventSink(&_xlApp.com());
            
            XLO_DEBUG(L"Made COM connection to Excel at '{}' with hwnd={}",
              (const wchar_t*)_xlApp.com().Path, (size_t)_excelWindowHandle);
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
          _xlApp.detach()->Release();
          CoUninitialize();
        }

        Application& excelApp() { return _xlApp; }
       
      private:
        HWND _excelWindowHandle;
        Application _xlApp;
        std::shared_ptr<Excel::AppEvents> _handler;
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
      auto result = theComConnector->excelApp().com().get_EnableEvents(&temp);
      return (SUCCEEDED(result));
    }

    Application& attachedApplication()
    {
      if (!theComConnector)
        throw ComConnectException("COM Connection not ready");
      return theComConnector->excelApp();
    }
  }
}