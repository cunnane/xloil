#include <xlOil/ExcelTypeLib.h>
#include <xlOil/AppObjects.h>
#include "ClassFactory.h"
#include "Connect.h"
#include <xlOil/Events.h>
#include <xlOil/ExcelThread.h>
#include <set>
#include <filesystem>

namespace fs = std::filesystem;
using std::set;
using std::wstring;
using std::pair;

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Maintains a list of open workbooks and signals the WorkbookAfterClose
    /// and WorkbookRename events. Excel's built-in WorkbookBeforeClose fires  
    /// before the user gets a chance to cancel the closure, so should not be
    /// used to run clean-up actions.
    /// </summary>
    class WorkbookMonitor
    {
    public:
      static void checkOnOpenWorkbook(struct Excel::_Workbook* Wb)
      {
        size_t numWorkbooks = thisApp().com().Workbooks->Count;

        // If workbook collection has grown by one, nothing was closed
        // and we just add the workbook name
        if (numWorkbooks == _workbooks.size() + 1)
          _workbooks.emplace(Wb->Name);
        else
          check();
      }

      /// <summary>
      /// On BeforeSave, we store the path of the active workbook
      /// </summary>
      static void Workbook_BeforeSave()
      {
        auto& app = thisApp().com();
        app.EnableEvents = false;
        _wbPathBeforeSave = fs::path((const wchar_t*)app.ActiveWorkbook->Path) 
          / (const wchar_t*)app.ActiveWorkbook->Name;
        app.EnableEvents = true;
      }

      /// <summary>
      /// On AfterSave we compare the path of the active workbook
      /// to the stored value and fire a WorkbookRename event if
      /// they differ
      /// </summary>
      static void Workbook_AfterSave(bool success)
      {
        if (!success)
          return;
        auto& app = thisApp().com();
        app.EnableEvents = false;
        const auto wbPath = fs::path((const wchar_t*)app.ActiveWorkbook->Path)
          / (const wchar_t*)app.ActiveWorkbook->Name;
        if (wbPath != _wbPathBeforeSave)
        {
          Event::WorkbookRename().fire(wbPath.c_str(), _wbPathBeforeSave.c_str());
        }
        app.EnableEvents = true;
      }

      static void check()
      {
        decltype(_workbooks) openWorkbooks;
        auto& app = thisApp().com();
        auto numWorkbooks = app.Workbooks->Count;
        for (auto i = 1; i <= numWorkbooks; ++i)
          openWorkbooks.emplace(app.Workbooks->Item[i]->FullName);

        std::vector<wstring> closedWorkbooks;
        std::set_difference(_workbooks.begin(), _workbooks.end(),
          openWorkbooks.begin(), openWorkbooks.end(), std::back_inserter(closedWorkbooks));

        for (auto& wb : closedWorkbooks)
        {
          auto lastSlash = wb.find_last_of(L'\\');
          Event::WorkbookAfterClose().fire(wb.c_str() + (lastSlash == wstring::npos ? 0 : lastSlash + 1));
        }

        _workbooks = openWorkbooks;
      }

      static set<wstring> _workbooks;

    private:
      static fs::path _wbPathBeforeSave;
    };

    set<wstring> WorkbookMonitor::_workbooks;
    fs::path WorkbookMonitor::_wbPathBeforeSave;

    constexpr auto DISPID_SheetSelectionChange = 0x616;
    constexpr auto DISPID_SheetActivate = 0x619;
    constexpr auto DISPID_SheetDeactivate = 0x61a;
    constexpr auto DISPID_SheetCalculate = 0x61b;
    constexpr auto DISPID_SheetChange = 0x61c;
    constexpr auto DISPID_AfterCalculate = 0xa34;

    class EventHandler :
      public ComEventHandler<NoIDispatchImpl<ComObject<Excel::AppEvents>>, Excel::AppEvents>
    {
    public:
      using Workbook = Excel::_Workbook;
      using Worksheet = Excel::_Worksheet;
      using Range = Excel::Range;

      EventHandler(Excel::_Application* source)
      {
        connect(source);
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
        if (inStack(DISPID_SheetSelectionChange))
          return;
        if (Event::SheetSelectionChange().handlers().empty())
          return;
        Event::SheetSelectionChange().fire(
          ((Worksheet*)Sh)->Name, ExcelRange(Target));
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
          ((Worksheet*)Sh)->Name, ExcelRange(Target), cancel);
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
          ((Worksheet*)Sh)->Name, ExcelRange(Target), cancel);
        *Cancel = cancel ? -1 : 0;
      }
      void SheetActivate(IDispatch* Sh)
      {
        if (inStack(DISPID_SheetActivate))
          return;
        Event::SheetActivate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetDeactivate(IDispatch* Sh)
      {
        if (inStack(DISPID_SheetDeactivate))
          return;
        Event::SheetDeactivate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetCalculate(IDispatch* Sh)
      {
        if (inStack(DISPID_SheetCalculate))
          return;
        if (Event::SheetCalculate().handlers().empty())
          return;
        Event::SheetCalculate().fire(((Worksheet*)Sh)->Name);
      }
      void SheetChange(
        IDispatch* Sh,
        Range* Target)
      {
        if (inStack(DISPID_SheetChange))
          return;
        if (Event::SheetChange().handlers().empty())
          return;
        Event::SheetChange().fire(
          ((Worksheet*)Sh)->Name, ExcelRange(Target));
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
        if (*Cancel == VARIANT_TRUE)
          return;
        bool cancel = false;
        Event::WorkbookBeforeClose().fire(Wb->Name, cancel);
        *Cancel = cancel ? VARIANT_TRUE : VARIANT_FALSE;
        if (!cancel)
        {
          // Wait 2s, then check if the workbook was actually closed. If the 
          // user still has the save/close dialog open, the COM call will fail
          // so we retry every 1 sec after that
          runExcelThread([]() { WorkbookMonitor::check(); }, ExcelRunQueue::COM_API, 2000, 1000);
        }
      }
      void WorkbookBeforeSave(
        Workbook* Wb,
        VARIANT_BOOL SaveAsUI,
        VARIANT_BOOL* Cancel)
      {
        bool cancel = *Cancel;
        Event::WorkbookBeforeSave().fire(Wb->Name, SaveAsUI < 0, cancel);
        WorkbookMonitor::Workbook_BeforeSave();
        *Cancel = cancel ? -1 : 0;
      }
      void WorkbookAfterSave(
        Workbook* Wb, 
        VARIANT_BOOL success)
      {
        Event::WorkbookAfterSave().fire(Wb->Name, success);
        WorkbookMonitor::Workbook_AfterSave(success);
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
      void AfterCalculate()
      {
        // Although we disable events, Excel turns them back on again
        // after calculation, so we can easily trigger a recursion by 
        // doing something in AfterCalculate which triggers a calculation.
        // We use this bool to protect against this.
        if (inStack(DISPID_AfterCalculate))
          return;

        Event::AfterCalculate().fire();
      }

      /// <summary>
      /// Checks if a given event dispid is already being processed. This isn't
      /// necessary for all events, just ones where their handlers could 
      /// re-trigger the event, e.g. if the SheetChange handler writes to the 
      /// sheet.
      /// </summary>>
      bool inStack(short dispid)
      {
        for (char i = 0; i < _stackCount - 1; ++i)
          if (_eventStack[i] == dispid)
            return true;
        return false;
      }

      STDMETHOD(Invoke)(DISPID dispidMember, REFIID /*riid*/,
        LCID /*lcid*/, WORD /*wFlags*/, DISPPARAMS* pdispparams, VARIANT* /*pvarResult*/,
        EXCEPINFO* pexcepinfo, UINT* /*puArgErr*/)
      {
        try
        {
          auto* rgvarg = pdispparams->rgvarg;

          // Note the event being processed. The event handling is single-threaded
          // but can be re-entrant. See comment for inStack
          _eventStack[_stackCount++] = (unsigned short) dispidMember;

          // These dispids are copied from oleview and are in the same order as listed there
          switch (dispidMember)
          {
          case 0x0000061d:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            NewWorkbook((Workbook*)rgvarg[0].pdispVal);
            break;
          case DISPID_SheetSelectionChange:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            SheetSelectionChange(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal);
            break;
          case 0x00000617:
            if (pdispparams->cArgs != 3) return DISP_E_BADPARAMCOUNT;
            SheetBeforeDoubleClick(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal, rgvarg[2].pboolVal);
            break;
          case 0x00000618:
            if (pdispparams->cArgs != 3) return DISP_E_BADPARAMCOUNT;
            SheetBeforeRightClick(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal, rgvarg[2].pboolVal);
            break;
          case DISPID_SheetActivate:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            SheetActivate(rgvarg[0].pdispVal);
            break;
          case DISPID_SheetDeactivate:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            SheetDeactivate(rgvarg[0].pdispVal);
            break;
          case DISPID_SheetCalculate:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            SheetCalculate(rgvarg[0].pdispVal);
            break;
          case DISPID_SheetChange:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            SheetChange(rgvarg[0].pdispVal, (Range*)rgvarg[1].pdispVal);
            break;
          case 0x0000061f:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            WorkbookOpen((Workbook*)rgvarg[0].pdispVal);
            break;
          case 0x00000620:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            WorkbookActivate((Workbook*)rgvarg[0].pdispVal);
            break;
          case 0x00000621:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            WorkbookDeactivate((Workbook*)rgvarg[0].pdispVal);
            break;
          case 0x00000622:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            WorkbookBeforeClose((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pboolVal);
            break;
          case 0x00000623:
            if (pdispparams->cArgs != 3) return DISP_E_BADPARAMCOUNT;
            WorkbookBeforeSave((Workbook*)rgvarg[0].pdispVal, rgvarg[1].boolVal, rgvarg[2].pboolVal);
            break;
          case 0x00000624:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            WorkbookBeforePrint((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pboolVal);
            break;
          case 0x00000625:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            WorkbookNewSheet((Workbook*)rgvarg[0].pdispVal, rgvarg[1].pdispVal);
            break;
          case 0x00000626:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            WorkbookAddinInstall((Workbook*)rgvarg[0].pdispVal);
            break;
          case 0x00000627:
            if (pdispparams->cArgs != 1) return DISP_E_BADPARAMCOUNT;
            WorkbookAddinUninstall((Workbook*)rgvarg[0].pdispVal);
            break;
          case DISPID_AfterCalculate:
            AfterCalculate();
            break;
          case 2911:
            if (pdispparams->cArgs != 2) return DISP_E_BADPARAMCOUNT;
            WorkbookAfterSave((Workbook*)rgvarg[0].pdispVal, rgvarg[1].boolVal);
            break;
          }
        }
        catch (_com_error& error)
        { 
          if (error.Error() == VBA_E_IGNORE)
          {
            pexcepinfo->scode = 0;
            pexcepinfo->wCode = error.WCode();
            pexcepinfo->bstrDescription = nullptr;
            pexcepinfo->bstrHelpFile = nullptr;
            pexcepinfo->bstrSource = nullptr;
            pexcepinfo->pfnDeferredFillIn = nullptr;
            return DISP_E_EXCEPTION;
          }
          XLO_ERROR(L"COM Error {0:#x}: {1}", (unsigned)error.Error(), error.ErrorMessage());
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Error during COM event handler callback: {0}", e.what());
        }
       
        // Pop the stack
        --_stackCount;

        return S_OK;
      }

    private:
      unsigned short _eventStack[16];
      char _stackCount = 0;
    };

    std::shared_ptr<Excel::AppEvents> createEventSink(Excel::_Application* source)
    {
      // We manage the COM object with a shared_ptr to avoid exporting ComPtr
      // everywhere. This means we need to AddRef/Release in the COM way.
      auto p = std::shared_ptr<EventHandler>(new EventHandler(source),
        [](auto* p)
        { 
          p->disconnect();
          p->Release(); 
        }); 
      p->AddRef();
      return p;
    }
    const std::set<std::wstring>& workbookPaths()
    {
      return WorkbookMonitor::_workbooks;
    }
  }
}