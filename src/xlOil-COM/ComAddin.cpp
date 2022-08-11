#include "ComAddin.h"
#include "ClassFactory.h"
#include "Connect.h"
#include "RibbonExtensibility.h"
#include "CustomTaskPane.h"
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/AppObjects.h>
#include <xlOil/State.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/ExcelUI.h>
#include <xlOil/Events.h>
#include <map>
#include <functional>

using std::wstring;
using std::map;
using std::vector;
using std::shared_ptr;
using std::unique_ptr;
using namespace Office;

namespace xloil
{
  namespace COM
  {
    struct ComAddinEvents
    {
      void OnDisconnection(bool /*excelClosing*/) {}
      void OnAddInsUpdate() { Event::ComAddinsUpdate().fire(); }
      void OnBeginShutdown() {}
    };

    class CustomTaskPaneConsumerImpl :
        public NoIDispatchImpl< ComObject<ICustomTaskPaneConsumer>>
    {
    public:
      HRESULT __stdcall raw_CTPFactoryAvailable(ICTPFactory* ctpfactory) override
      {
        factory = ctpfactory;
        return S_OK;
      }

      STDMETHOD(QueryInterface)(REFIID riid, void** ppv) noexcept
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == IID_IDispatch
          || riid == __uuidof(ICustomTaskPaneConsumer))
        {
          *ppv = this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }

      STDMETHOD(Invoke)(
        _In_ DISPID dispidMember,
        _In_ REFIID /*riid*/,
        _In_ LCID /*lcid*/,
        _In_ WORD /*wFlags*/,
        _In_ DISPPARAMS* pdispparams,
        _Out_opt_ VARIANT* /*pvarResult*/,
        _Out_opt_ EXCEPINFO* /*pexcepinfo*/,
        _Out_opt_ UINT* /*puArgErr*/) override
      {
        // Remember the args are in reverse order
        auto* rgvarg = pdispparams->rgvarg;

        if (dispidMember == 1)
        {
          return raw_CTPFactoryAvailable((Office::ICTPFactory*)rgvarg[0].pdispVal);
        }

        XLO_ERROR("Internal Error: unknown dispid called on task pane consumer Invoke.");
        return E_FAIL;
      }

      ICTPFactory* factory = nullptr;
    };

    // This class does not need a disp-interface
    class ComAddinImpl :
        public NoIDispatchImpl<ComObject<AddInDesignerObjects::IDTExtensibility2>>
    {
    public:
      ComAddinImpl()
        : _customTaskPane(new CustomTaskPaneConsumerImpl())
        , ribbon(nullptr)
      {}

      STDMETHOD(QueryInterface)(REFIID riid, void** ppv) noexcept
      {
        *ppv = NULL;
        if (riid == IID_IUnknown 
          || riid == __uuidof(AddInDesignerObjects::IDTExtensibility2))
        {
          auto p = (AddInDesignerObjects::IDTExtensibility2*)this;
          *ppv = p;
          p->AddRef();
          return S_OK;
        }
        else if (riid == __uuidof(IRibbonExtensibility))
        {
          if (ribbon)
            return ribbon->QueryInterface(riid, ppv);
        }
        else if (riid == __uuidof(ICustomTaskPaneConsumer))
        {
          return _customTaskPane->QueryInterface(riid, ppv);
        }
        return E_NOINTERFACE;
      }
      virtual HRESULT __stdcall raw_OnConnection(
        /*[in]*/ IDispatch * /*Application*/,
        /*[in]*/ enum AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/,
        /*[in]*/ IDispatch * /*AddInInst*/,
        /*[in]*/ SAFEARRAY**) override
      {
        return S_OK;
      }
      virtual HRESULT __stdcall raw_OnDisconnection(
        /*[in]*/ enum AddInDesignerObjects::ext_DisconnectMode RemoveMode,
        /*[in]*/ SAFEARRAY**) override
      {
        events->OnDisconnection(
          RemoveMode == AddInDesignerObjects::ext_DisconnectMode::ext_dm_HostShutdown);
        return S_OK;
      }
      virtual HRESULT __stdcall raw_OnAddInsUpdate(SAFEARRAY**) override
      {
        events->OnAddInsUpdate();
        return S_OK;
      }
      virtual HRESULT __stdcall raw_OnStartupComplete(SAFEARRAY**) override
      {
        return S_OK;
      }
      virtual HRESULT __stdcall raw_OnBeginShutdown(SAFEARRAY**) override
      {
        events->OnBeginShutdown();
        return S_OK;
      }

      ICTPFactory* ctpFactory() const
      {
        return _customTaskPane ? _customTaskPane->factory : nullptr;
      }

      IRibbonExtensibility* ribbon;
      CComPtr<CustomTaskPaneConsumerImpl> _customTaskPane;
      unique_ptr<ComAddinEvents> events;
    };

    struct SetAutomationSecurity
    {
      SetAutomationSecurity(Office::MsoAutomationSecurity value)
      {
        _previous = excelApp().com().AutomationSecurity;
        excelApp().com().AutomationSecurity = value;
      }
      ~SetAutomationSecurity()
      {
        try
        {
          excelApp().com().AutomationSecurity = _previous;
        }
        catch (...)
        {
        }
      }
      Office::MsoAutomationSecurity _previous;
    };

    class ComAddinCreator : public IComAddin
    {
    private:
      auto& comAddinImpl() const
      {
        return _registrar.server();
      }

      RegisterCom<ComAddinImpl> _registrar;
      bool                      _connected = false;
      shared_ptr<IRibbon>       _ribbon;
      COMAddIn*                 _comAddin = nullptr;
      TaskPaneMap               _panes;
      shared_ptr<const void>    _closeHandler;
    
      ComAddinCreator(const wchar_t* name, const wchar_t* description)
        : _registrar(
          [](const wchar_t*, const GUID&) { return new ComAddinImpl(); },
          formatStr(L"%s.ComAddin", name ? name : L"xlOil").c_str())
      {
        // TODO: hook OnDisconnect to stop user from disabling COM stub.
        comAddinImpl()->events.reset(new ComAddinEvents());

        if (!name)
          XLO_THROW("Com add-in name must be provided");

        // It's possible the addin has already been registered and loaded and 
        // is just being reinitialised, so we do findAddin twice
        auto& app = excelApp().com();

        SetAutomationSecurity setSecurity(
          Office::MsoAutomationSecurity::msoAutomationSecurityLow);

        findAddin(app);

        if (isComAddinConnected())
        {
          disconnect();
        }
        else
        {
          const auto addinPath = fmt::format(
            L"Software\\Microsoft\\Office\\Excel\\AddIns\\{0}", _registrar.progid());
          _registrar.writeRegistry(
            HKEY_CURRENT_USER, addinPath.c_str(), L"FriendlyName", name);
          _registrar.writeRegistry(
            HKEY_CURRENT_USER, addinPath.c_str(), L"LoadBehavior", (DWORD)0);
          if (description)
            _registrar.writeRegistry(
              HKEY_CURRENT_USER, addinPath.c_str(), L"Description", description);

          app.GetCOMAddIns()->Update();
          findAddin(app);
          if (!_comAddin)
            XLO_THROW(L"Add-in connect: could not find addin '{0}'", progid());
        }
      }

    public:
      static auto create(const wchar_t* name, const wchar_t* description)
      {
        auto p = std::shared_ptr<ComAddinCreator>(new ComAddinCreator(name, description));
        p->_closeHandler = Event::WorkbookAfterClose().weakBind(
          std::weak_ptr<ComAddinCreator>(p), 
          &ComAddinCreator::handleWorkbookClose);
        return p;
      }

      ~ComAddinCreator()
      {
        try
        {
          _closeHandler.reset();
          for (auto& pane : _panes)
            pane.second->destroy();

          _panes.clear();
          disconnect();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("ComAddin failed to close: {0}", e.what());
        }
      }

      void findAddin(Excel::_Application& app)
      {
        auto ourProgid = _variant_t(progid());
        app.GetCOMAddIns()->raw_Item(&ourProgid, &_comAddin); // TODO: Does this need decref/incref?
      }

      bool isComAddinConnected() const
      {
        try
        {
          return _comAddin
            ? (_comAddin->Connect == VARIANT_TRUE)
            : false;
        }
        XLO_RETHROW_COM_ERROR;
      }

      void connect(
        const wchar_t* xml,
        const RibbonMap& mapper) override
      {
        if (_connected)
          return;
        try
        {
          if (xml)
          {
            _ribbon = createRibbon(xml, mapper);
            comAddinImpl()->ribbon = _ribbon->getRibbon();
          }
          _comAddin->Connect = VARIANT_TRUE;
          _connected = true;
        }
        catch (_com_error& error)
        {
          if (error.Error() == 0x80004004) // Operation aborted
          {
            XLO_THROW("During add-in connect, received Operation Aborted ({}) this probably indicates "
              "blocking by add-in security.  Check add-ins are enabled and this add-in is not a disabled "
              "COM add-in.", (unsigned)error.Error());
          }
          else
            XLO_THROW(L"COM Error {0:#x}: {1}", (unsigned)error.Error(), error.ErrorMessage()); \
        }
      }

      // TODO: use of disconnect is fatal to any attached custom task panes, they do not
      // reappear on connect
      void disconnect() override
      {
        if (!_connected)
          return;
        try
        {
          _comAddin->Connect = VARIANT_FALSE;
          _connected = false;
        }
        XLO_RETHROW_COM_ERROR;
      }

      const wchar_t* progid() const override
      {
        return _registrar.progid();
      }
      void ribbonInvalidate(const wchar_t* controlId = 0) const override
      {
        if (_ribbon)
          _ribbon->invalidate(controlId);
      }
      bool ribbonActivate(const wchar_t* controlId) const override
      {
        return _ribbon
          ? _ribbon->activateTab(controlId)
          : false;
      }

      shared_ptr<ICustomTaskPane> createTaskPane(
        const wchar_t* name,
        const ExcelWindow* window,
        const wchar_t* progId) override
      {
        auto factory = comAddinImpl()->ctpFactory();
        if (!factory)
          XLO_THROW("Internal error: failed to receive CTP factory");

        shared_ptr<ICustomTaskPane> pane(
          createCustomTaskPane(
            *factory, 
            name, 
            window ? window->dispatchPtr() : nullptr, 
            progId));

        _panes.insert(make_pair(pane->window().workbook().name(), pane));

        return pane;
      }

      const TaskPaneMap& panes() const override { return _panes; }

      void handleWorkbookClose(const wchar_t* wbName)
      {
        // destroy
        _panes.erase(wbName);
      }
    };

    shared_ptr<IComAddin> createComAddin(
      const wchar_t* name, const wchar_t* description)
    {
      return ComAddinCreator::create(name, description);
    }
  }
}