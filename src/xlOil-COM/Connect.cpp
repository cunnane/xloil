#include "Connect.h"
#include "ComEventSink.h"
#include <xloil/AppObjects.h>
#include <xloil/Events.h>
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/State.h>
#include <oleacc.h> // AccessibleObjectFromWindow
#include <memory>

using std::make_shared;
using std::wstring;

namespace xloil
{
  namespace COM
  {
    HWND nextExcelMainWindow(HWND startFrom)
    {
      return FindWindowExA(0, startFrom, "XLMAIN", NULL);
    }

    Excel::_Application* newApplicationObject()
    {
      CoInitialize(NULL); // It's safe to call repeatedly
      Excel::_ApplicationPtr app;
      auto hr = app.CreateInstance(
        __uuidof(Excel::Application),
        NULL,
        CLSCTX_LOCAL_SERVER);
      if (hr == S_OK)
      {
          return app.Detach();
      }
      else
      {
          _com_error error(hr);
          throw ComConnectException(
              utf16ToUtf8(
                  fmt::format(L"Failed to create Application object. COM Error {0:#x}: {1}",
                      (size_t)error.Error(), error.ErrorMessage())).c_str());
      }
    }

    Excel::_Application* applicationObjectFromWindow(HWND xlmainHandle)
    {
      CoInitialize(NULL);
      
      // Based on discussion https://stackoverflow.com/questions/30363748/
      // However we need to make a modification to the method given there:
      // 
      // The result of `pWindow->Application` differs from 
      // `pWindow->Parent->Application` if we *do not own* the Excel process
      // (i.e. we didn't start it and are not running within it). In this case
      // the window's Application object is broken and will dump core when used.
      // The only difference I can see with the VBA implementation is that VBA
      // uses late binding, i.e. calls `invoke` which may take a different code 
      // path.  This does look awfully like a bug, but since it's in the COM 
      // interface its probably existed for years and will never be fixed.
      //
      auto hwnd2 = FindWindowExA(xlmainHandle, 0, "XLDESK", NULL);
      auto hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", NULL);
      Excel::Window* pWindow = NULL;
      HRESULT hr = AccessibleObjectFromWindow(hwnd3, (DWORD)OBJID_NATIVEOM,
          __uuidof(Excel::Window),
          (void**)&pWindow);

      if ( hr == S_OK)
      {
        auto parent = pWindow->Parent;
        Excel::_WorkbookPtr parentWorkbook(parent);
        pWindow->Release();
        
        auto result = parentWorkbook->Application;
        return result.Detach();
      }
      else
      {
        _com_error error(hr);
        throw ComConnectException(
            utf16ToUtf8(
                fmt::format(L"Failed to create Application object from window handle. COM Error {0:#x}: {1}. {2}",
                    (size_t)error.Error(), error.ErrorMessage(), error.Description())).c_str());
      }
      return nullptr;
    }

    namespace
    {
      class COMConnector
      {
      public:
        COMConnector()
          : _excelWindowHandle((HWND)Environment::excelProcess().hWnd)
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
          _xlApp.release();
          CoUninitialize();
        }

        Application& thisApp() { return _xlApp; }
       
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
      auto result = theComConnector->thisApp().com().get_EnableEvents(&temp);
      return (SUCCEEDED(result));
    }

    Application& attachedApplication()
    {
      if (!theComConnector)
        throw ComConnectException("COM Connection not ready");
      return theComConnector->thisApp();
    }
  }
}