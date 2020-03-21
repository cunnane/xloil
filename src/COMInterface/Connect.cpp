
#include "Connect.h"

#include "xloil/Events.h"
#include "xloil/ExcelCall.h"
#include <oleacc.h> // must include before ExcelTypeLib
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

template<class TSource>
void connect(TSource* source, IDispatch* sink, IConnectionPoint*& connPoint, DWORD& eventCookie)
{
  IConnectionPointContainer *pConnPntCont;
  IUnknown* pIUnknown = NULL;

  sink->QueryInterface(IID_IUnknown, (void**)(&pIUnknown));
  source->QueryInterface(IID_IConnectionPointContainer, (void**)&pConnPntCont);

  if (pConnPntCont)
  {
    pConnPntCont->FindConnectionPoint(__uuidof(Excel::AppEvents), &connPoint);
    pConnPntCont->Release();
    pConnPntCont = NULL;
  }

  if (connPoint)
    connPoint->Advise(pIUnknown, &eventCookie);

  pIUnknown->Release();
}

namespace xloil
{
  class WorkbookMonitor
  {
  public:
    static void checkOnOpenWorkbook(struct Excel::_Workbook* Wb)
    {
      auto numWorkbooks = excelApp().Workbooks->Count;

      // If workbook collection has grown by one, nothing was closed
      // and we just add the workbook name
      if (numWorkbooks == _workbooks.size() + 1)
        _workbooks.emplace(Wb->Name);
      else
        check();
    }
    static void check()
    {
      set<wstring> workbooks;
      auto& app = excelApp();
      auto numWorkbooks = app.Workbooks->Count;
      for (auto i = 1; i <= numWorkbooks; ++i)
        workbooks.emplace(app.Workbooks->Item[i]->Name);

      std::vector<wstring> closedWorkbooks;
      std::set_difference(_workbooks.begin(), _workbooks.end(),
        workbooks.begin(), workbooks.end(), std::back_inserter(closedWorkbooks));

      for (auto& wb : closedWorkbooks)
        Event_WorkbookClose().fire(wb.c_str());

      _workbooks = workbooks;
    }
  private:
    static set<wstring> _workbooks;
  };

  set<wstring> WorkbookMonitor::_workbooks;
}

class EventHandler : Excel::AppEvents
{
public:
  template <class T> 
  EventHandler(T* source)
    : _cRef(1)
  {
    connect(source, this, _pIConnectionPoint, _dwEventCookie);
  }

  ~EventHandler()
  {
    close();
  }

  void close()
  {
    if (_pIConnectionPoint)
    {
      _pIConnectionPoint->Unadvise(_dwEventCookie);
      _dwEventCookie = 0;
      _pIConnectionPoint->Release();
      _pIConnectionPoint = NULL;
    }
  }

  HRESULT NewWorkbook(struct Excel::_Workbook* Wb)
  {
    xloil::Event_NewWorkbook().fire(Wb->Name);
    xloil::WorkbookMonitor::checkOnOpenWorkbook(Wb);
    return 0;
  }
  HRESULT WorkbookOpen(struct Excel::_Workbook* Wb)
  {
    xloil::Event_WorkbookOpen().fire(Wb->Path, Wb->Name);
    xloil::WorkbookMonitor::checkOnOpenWorkbook(Wb);
    return 0;
  }
  HRESULT SheetActivate(IDispatch * Sh)
  {
    return 0;
  }

  STDMETHOD_(ULONG, AddRef)()
  {
    InterlockedIncrement(&_cRef);
    return _cRef;
  }

  STDMETHOD_(ULONG, Release)()
  {
    InterlockedDecrement(&_cRef);
    if (_cRef == 0)
    {
      delete this;
      return 0;
    }
    return _cRef;
  }

  STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObject)
  {
    if (riid == IID_IUnknown)
    {
      *ppvObject = (IUnknown*)this;
      AddRef();
      return S_OK;
    }
    else if ((riid == IID_IDispatch) || (riid == __uuidof(Excel::AppEvents)))
    {
      *ppvObject = (IDispatch*)this;
      AddRef();
      return S_OK;
    }

    return E_NOINTERFACE;
  }

  STDMETHOD(GetTypeInfoCount)(UINT* pctinfo)
  {
    return E_NOTIMPL;
  }

  STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
  {
    return E_NOTIMPL;
  }

  STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
    LCID lcid, DISPID* rgdispid)
  {
    return E_NOTIMPL;
  }

  STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid,
    LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,
    EXCEPINFO* pexcepinfo, UINT* puArgErr)
  {
    if ((riid != IID_NULL))
      return E_INVALIDARG;

    // Note for adding handlers: the rgvarg array is backwards
    switch (dispidMember) 
    {
    case 0x0000061d:
      NewWorkbook((Excel::_Workbook*)pdispparams->rgvarg[0].pdispVal);
      break;
    case 0x0000061f:
      WorkbookOpen((Excel::_Workbook*)pdispparams->rgvarg[0].pdispVal);
      break;
    case 0x00000619:
      //SheetActivate((IDispatch*)pdispparams->rgvarg[0].pdispVal);
      break;
    }

    return S_OK;
  }
private:
  IConnectionPoint* _pIConnectionPoint;
  DWORD	_dwEventCookie;
  LONG _cRef;
};

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
          _handler.reset(new EventHandler(p));
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
      std::unique_ptr<EventHandler> _handler;
      HWND _excelWindowHandle;
    };


    struct RegisterMe
    {
      RegisterMe()
      {
        connector = nullptr;
        static auto handler = xloil::Event_AutoClose() += [this]() { delete connector; };
      }
      COMConnector* connector;
      COMConnector* connect()
      {
        if (!connector)
          connector = new COMConnector();
        return connector;
      }
    } theInstance;
  }

  void reconnectCOM()
  {
    if (theInstance.connector)
    {
      delete theInstance.connector;
      theInstance.connector = nullptr;
    }
    theInstance.connect();
  }

  Excel::_Application& excelApp()
  {
    return *theInstance.connect()->ExcelApp();
  }
}