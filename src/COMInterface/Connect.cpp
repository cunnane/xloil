
#include "Connect.h"
#include <oleacc.h>
#include "xloil/Events.h"
#include "ExcelObj.h"
#include "xloil/Register.h"
#include "xloil/ExcelCall.h"
#include <memory>
#include "ExcelTypeLib.h"
#include <unordered_set>

using std::make_shared;
using std::wstring;
using std::unordered_set;

/// <summary>
/// A naive GetActiveObject("Excel.Application") gets the first instance of Excel
/// </summary>
Excel::_ApplicationPtr getExcelInstance(HWND xlmainHandle)
{
  // Based on discussion here:
  // https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
  HWND hwnd = nullptr, hwnd2, hwnd3;
  hwnd2 = FindWindowExA(xlmainHandle, 0, "XLDESK", NULL);
  hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", NULL);
  Excel::Window* pWindow = NULL;

  // Sometimes AccessibleObjectFromWindow fails for no apparent reason. Retry
  for (auto tries = 0; tries < 10; ++tries)
  {
    if (AccessibleObjectFromWindow(hwnd3, OBJID_NATIVEOM, __uuidof(IDispatch), (void**)&pWindow) == S_OK)
      return pWindow->Application;
    Sleep(100);
  }
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
      unordered_set<wstring> workbooks;
      auto numWorkbooks = excelApp().Workbooks->Count;
      for (auto i = 1; i <= numWorkbooks; ++i)
        workbooks.emplace(excelApp().Workbooks->Item[i]->Name);

      std::vector<wstring> closedWorkbooks;
      std::set_difference(_workbooks.begin(), _workbooks.end(),
        workbooks.begin(), workbooks.end(), std::back_inserter(closedWorkbooks));

      for (auto& wb : closedWorkbooks)
        Event_WorkbookClose().fire(wb.c_str());

      _workbooks = workbooks;
    }
  private:
    static unordered_set<wstring> _workbooks;
  };

  unordered_set<wstring> WorkbookMonitor::_workbooks;
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


class COMConnector
{
public:
  COMConnector()
  {
    try
    {
      CoInitialize(NULL);
      auto windowsHandle = xloil::callExcel(msxll::xlGetHwnd);
      // This conversion is OK even in x64 because the window handle is an index
      // into an array, not a pointer. 
#pragma warning(disable: 4312)
      XL = getExcelInstance((HWND)windowsHandle.toInt());
      
      Excel::_Application* p = XL;
      _handler.reset(new EventHandler(p));
    }
    catch (_com_error& error)
    {
      XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }

  ~COMConnector()
  {
    CoUninitialize();
  }

  Excel::_ApplicationPtr& ExcelApp() { return XL; }

private:
  Excel::_ApplicationPtr XL;
  std::unique_ptr<EventHandler> _handler;
};

// See https://social.msdn.microsoft.com/Forums/vstudio/en-US/9168f9f2-e5bc-4535-8d7d-4e374ab8ff09/hresult-800ac472-from-set-operations-in-excel?forum=vsto
constexpr HRESULT VBA_E_IGNORE = 0x800ac472;

template <class TFunc>
bool retryComCall(TFunc fn)
{
  XLO_TRACE("Calling into XLL context fn= {0:#x}", (size_t)&fn);
  for (auto tries = 0; tries < 10; ++tries)
  {
    try
    {
      fn();
      return true;
    }
    catch (_com_error& error)
    {
      if (error.Error() != VBA_E_IGNORE)
      {
        XLO_ERROR(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
        break;
      }
    }
    Sleep(50);
    XLO_TRACE("Retry # {0} for COM call", (tries + 1));
  }
  return false;
}
namespace xloil
{
  Excel::_Application& excelApp()
  {
    static COMConnector c;
    return c.ExcelApp();
  }

  static const std::function<void()>* theTargetFunc = nullptr;

  // TODO: make these commmands so they are hidden and have void return?
  XLO_ENTRY_POINT(XLOIL_XLOPER*) xloRunFuncInXLLContext()
  {
    // Do we need this result?
    static ExcelObj result;
    try
    {
      ScopeInXllContext context;
      (*theTargetFunc)();
    }
    catch (...)
    {
    }
    return &result;
  }
  XLO_REGISTER(xloRunFuncInXLLContext)
    .macro();

  static int theExcelCallFunc = 0;
  static XLOIL_XLOPER* theExcelCallResult = nullptr;
  static XLOIL_XLOPER** theExcelCallArgs = nullptr;
  static int theExcelCallNumArgs = 0;

  XLO_ENTRY_POINT(XLOIL_XLOPER*) xloRunInXLLContext()
  {
    static ExcelObj result;
    try
    {
      ScopeInXllContext context;
      Excel12v(theExcelCallFunc, theExcelCallResult, theExcelCallNumArgs, theExcelCallArgs);
    }
    catch (...)
    {
    }
    return &result;
  }
  XLO_REGISTER(xloRunInXLLContext)
    .macro();
  
  ScopeInXllContext::ScopeInXllContext()
  {
    ++_count;
  }
  ScopeInXllContext::~ScopeInXllContext()
  {
    --_count;
  }
  bool ScopeInXllContext::check() 
  { 
    return _count > 0; 
  }

  int ScopeInXllContext::_count = 0;



  bool runInXllContext(const std::function<void()>& f)
  {
    if (ScopeInXllContext::check())
    {
      f();
      return true;
    }
   
    auto[result, xlret] = tryCallExcel(msxll::xlfGetDocument, 1);
    if (xlret == 0)
    {
      f();
      return true;
    }

    theTargetFunc = &f;

    return retryComCall([]() { excelApp().Run("xloRunFuncInXLLContext"); });
  }

  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args)
  {
    if (ScopeInXllContext::check())
    {
      Excel12v(func, result, nArgs, (XLOIL_XLOPER**)args);
      return true;
    }
    theExcelCallFunc = func;
    theExcelCallResult = result;
    theExcelCallArgs = (XLOIL_XLOPER**)args;
    theExcelCallNumArgs = nArgs;
    return retryComCall([]() { excelApp().Run("xloRunInXLLContext"); });
  }
}