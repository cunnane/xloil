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
  using detail::AppObject;

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
      return hr == S_OK ? app.Detach() : nullptr;
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
      if (AccessibleObjectFromWindow(hwnd3, (DWORD)OBJID_NATIVEOM,
        __uuidof(Excel::Window),
        (void**)&pWindow) == S_OK)
      {
        auto parent = pWindow->Parent;
        Excel::_WorkbookPtr parentWorkbook(parent);
        pWindow->Release();
        
        auto result = parentWorkbook->Application;
        return result.Detach();
      }
      return nullptr;
    }

    namespace
    {
      class COMConnector
      {
      public:
        COMConnector()
          : _xlApp(nullptr)
        {}

        bool connect()
        {
          if (connected())
            return true;

          try
          {
            auto excelWindowHandle = (HWND)Environment::excelProcess().hWnd;
            auto application = applicationObjectFromWindow(excelWindowHandle);
            if (!application)
              return false;

            _xlApp = Application(application);
            
            XLO_DEBUG(L"Made COM connection to Excel at '{}' with hwnd={}",
              (const wchar_t*)_xlApp.com().Path, (size_t)excelWindowHandle);

            _handler = COM::createEventSink(&_xlApp.com());
          }
          catch (_com_error& error)
          {
            XLO_DEBUG(L"COM connect failed: {0:#x}: {1}",
              (size_t)error.Error(), error.ErrorMessage());
          }

          return connected();
        }

        bool connected() const { return _xlApp.valid(); }

        ~COMConnector()
        {
          _handler.reset();
          _xlApp.release();
          CoUninitialize();
        }

        Application& thisApp() { return (Application&)_xlApp; }
       
      private:
        AppObject<Excel::_Application, true> _xlApp;
        std::shared_ptr<Excel::AppEvents> _handler;
      };
    }

    static std::unique_ptr<COMConnector> theComConnector(new COMConnector());

    bool connectCom()
    {
      if (theComConnector->connected())
        return true;
      else
        return theComConnector->connect();
    }

    void disconnectCom()
    {
      theComConnector.reset();
    }

    bool isComApiAvailable() noexcept
    {
      if (!theComConnector->connected())
        return false;
      
      // Do some random COM thing - is this the fastest thing?
      VARIANT_BOOL temp;
      auto result = theComConnector->thisApp().com().get_EnableEvents(&temp);
      return (SUCCEEDED(result));
    }

    Application& attachedApplication()
    {
      if (!theComConnector->connected())
        throw ComConnectException("COM Connection not ready");
      return theComConnector->thisApp();
    }
  }
}