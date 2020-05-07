#include <xlOil/Events.h>
#include <oleacc.h> // must include before ExcelTypeLib
#include "ExcelTypeLib.h"
#include "Connect.h"
#include "ComRange.h"
#include "XllContextInvoke.h"
#include <set>

using std::set;
using std::wstring;

namespace xloil
{
  namespace COM
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
          Event::WorkbookAfterClose().fire(wb.c_str());

        _workbooks = workbooks;
      }
    private:
      static set<wstring> _workbooks;
    };

    set<wstring> WorkbookMonitor::_workbooks;

    template<class TSource>
    void connectSourceToSink(
      TSource* source, 
      IDispatch* sink,
      IConnectionPoint*& connectionPoint, 
      DWORD& eventCookie)
    {
      IConnectionPointContainer *pContainter;
      IUnknown* pIUnknown = nullptr;

      // Get IUnknown for sink
      sink->QueryInterface(IID_IUnknown, (void**)(&pIUnknown));

      // Get connection point for source
      source->QueryInterface(IID_IConnectionPointContainer, (void**)&pContainter);
      if (pContainter)
      {
        pContainter->FindConnectionPoint(__uuidof(Excel::AppEvents), &connectionPoint);
        pContainter->Release();
      }

      if (connectionPoint)
        connectionPoint->Advise(pIUnknown, &eventCookie);

      pIUnknown->Release();
    }

    class EventHandler : public Excel::AppEvents
    {
    public:
      using Workbook = Excel::_Workbook;
      using Worksheet = Excel::_Worksheet;
      using Range = Excel::Range;

      EventHandler(Excel::_Application* source)
        : _cRef(1)
      {
        connectSourceToSink(source, this, _pIConnectionPoint, _dwEventCookie);
      }

      virtual ~EventHandler()
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

      void NewWorkbook(Workbook* Wb)
      {
        Event::NewWorkbook().fire(Wb->Name);
        WorkbookMonitor::checkOnOpenWorkbook(Wb);
      }
      void SheetSelectionChange(
        IDispatch* Sh,
        Range* Target)
      {
        if (Event::SheetSelectionChange().handlers().empty())
          return;
        Event::SheetSelectionChange().fire(
          ((Worksheet*)Sh)->Name, rangeFromComRange(Target));
      }
      void SheetBeforeDoubleClick(
        IDispatch* Sh,
        Range* Target,
        VARIANT_BOOL* Cancel)
      {
        if (Event::SheetBeforeDoubleClick().handlers().empty())
          return;

        bool cancel = *Cancel;
        Event::SheetBeforeDoubleClick().fire(
          ((Worksheet*)Sh)->Name, rangeFromComRange(Target), cancel);
        *Cancel = cancel ? -1 : 0;
      }
      void SheetBeforeRightClick(
        IDispatch* Sh,
        Range* Target,
        VARIANT_BOOL* Cancel)
      {
        if (Event::SheetBeforeRightClick().handlers().empty())
          return;

        bool cancel = *Cancel;
        Event::SheetBeforeRightClick().fire(
          ((Worksheet*)Sh)->Name, rangeFromComRange(Target), cancel);
        *Cancel = cancel ? -1 : 0;
      }
      void SheetActivate(IDispatch* Sh)
      {
        Event::SheetActivate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetDeactivate(IDispatch* Sh)
      {
        Event::SheetDeactivate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetCalculate(IDispatch* Sh)
      {
        if (Event::SheetCalculate().handlers().empty())
          return;
        Event::SheetCalculate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetChange(
        IDispatch* Sh,
        Range* Target)
      {
        if (Event::SheetChange().handlers().empty())
          return;
        Event::SheetChange().fire(
          ((Worksheet*)Sh)->Name, rangeFromComRange(Target));
      }
      void WorkbookOpen(Workbook* Wb)
      {
        Event::WorkbookOpen().fire(Wb->Path, Wb->Name);
        WorkbookMonitor::checkOnOpenWorkbook(Wb);
      }
      void WorkbookActivate(Workbook* Wb)
      {
        Event::WorkbookActivate().fire(Wb->Name);
      }
      void WorkbookDeactivate(Workbook* Wb)
      {
        Event::WorkbookDeactivate().fire(Wb->Name);
      }
      void WorkbookBeforeClose(
        Workbook* Wb,
        VARIANT_BOOL* Cancel)
      {
        //bool cancel = *Cancel;
        //Event::WorkbookBeforeClose().fire(Wb->Name, cancel);
        //*Cancel = cancel ? -1 : 0;
      }
      void WorkbookBeforeSave(
        Workbook* Wb,
        VARIANT_BOOL SaveAsUI,
        VARIANT_BOOL* Cancel)
      {
        bool cancel = *Cancel;
        Event::WorkbookBeforeSave().fire(Wb->Name, SaveAsUI < 0, cancel);
        *Cancel = cancel ? -1 : 0;
      }
      void WorkbookBeforePrint(
        Workbook* Wb,
        VARIANT_BOOL* Cancel)
      {
        bool cancel = *Cancel;
        Event::WorkbookBeforePrint().fire(Wb->Name, cancel);
        *Cancel = cancel ? -1 : 0;
      }
      void WorkbookNewSheet(
        Workbook* Wb,
        IDispatch* Sh)
      {
        Event::WorkbookNewSheet().fire(Wb->Name, ((Worksheet*)Sh)->Name);
      }
      void WorkbookAddinInstall(Workbook* Wb)
      {
        Event::WorkbookAddinInstall().fire(Wb->Name);
      }
      void WorkbookAddinUninstall(Workbook* Wb)
      {
        Event::WorkbookAddinUninstall().fire(Wb->Name);
      }

      STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid,
        LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,
        EXCEPINFO* pexcepinfo, UINT* puArgErr)
      {
        if ((riid != IID_NULL))
          return E_INVALIDARG;

        // TODO: TRY CATCH
        auto* rgvarg = pdispparams->rgvarg;

        // These dispids are copied from oleview and are in the same
        // order as listed there

        InComContext scope;

        switch (dispidMember)
        {
        case 0x0000061d:
          NewWorkbook((Workbook*)rgvarg[0].pdispVal);
          break;
        case 0x00000616:
          SheetSelectionChange(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal);
          break;
        case 0x00000617:
          SheetBeforeDoubleClick(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal, rgvarg[2].pboolVal);
          break;
        case 0x00000618:
          SheetBeforeRightClick(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal, rgvarg[2].pboolVal);
          break;
        case 0x00000619:
          SheetActivate(rgvarg[0].pdispVal);
          break;
        case 0x0000061a:
          SheetDeactivate(rgvarg[0].pdispVal);
          break;
        case 0x0000061b:
          SheetCalculate(rgvarg[0].pdispVal);
          break;
        case 0x0000061c:
          SheetChange(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal);
          break;
        case 0x0000061f:
          WorkbookOpen((Workbook*)rgvarg[0].pdispVal);
          break;
        case 0x00000620:
          WorkbookActivate((Workbook*)rgvarg[0].pdispVal);
          break;
        case 0x00000621:
          WorkbookDeactivate((Workbook*)rgvarg[0].pdispVal);
          break;
        case 0x00000622:
          WorkbookBeforeClose((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pboolVal);
          break;
        case 0x00000623:
          WorkbookBeforeSave((Workbook*)rgvarg[0].pdispVal, rgvarg[1].boolVal, rgvarg[2].pboolVal);
          break;
        case 0x00000624:
          WorkbookBeforePrint((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pboolVal);
          break;
        case 0x00000625:
          WorkbookNewSheet((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pdispVal);
          break;
        case 0x00000626:
          WorkbookAddinInstall((Workbook*)rgvarg[0].pdispVal);
          break;
        case 0x00000627:
          WorkbookAddinUninstall((Workbook*)rgvarg[0].pdispVal);
          break;
        }

        return S_OK;
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

    private:
      IConnectionPoint* _pIConnectionPoint;
      DWORD	_dwEventCookie;
      LONG _cRef;
    };

    std::shared_ptr<Excel::AppEvents> createEventSink(Excel::_Application* source)
    {
      return std::make_shared<EventHandler>(source);
    }
  }
}