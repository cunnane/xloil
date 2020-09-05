#include "ComAddin.h"
#include "ExcelTypeLib.h"
#include "ClassFactory.h"
#include "Connect.h"
#include "RibbonExtensibility.h"
#include <xlOil/State.h>
#include <xlOil/Log.h>
#include <xlOil/Interface.h>
#include <map>
#include <functional>

using std::wstring;
using std::map;
using std::vector;

namespace xloil
{
  namespace COM
  {
    class ComAddinCreator : public IComAddin
    {
      // This class does not need a disp-interface
      class __declspec(novtable)
        ComAddinImpl :
          public CComObjectRootEx<CComSingleThreadModel>,
          public NoIDispatchImpl<AddInDesignerObjects::IDTExtensibility2>
      {
      public:
        HRESULT _InternalQueryInterface(REFIID riid, void** ppv) throw()
        {
          *ppv = NULL;
          if (riid == IID_IUnknown || riid == __uuidof(AddInDesignerObjects::IDTExtensibility2))
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
          return E_NOINTERFACE;
        }
        virtual HRESULT __stdcall raw_OnConnection(
          /*[in]*/ IDispatch * Application,
          /*[in]*/ enum AddInDesignerObjects::ext_ConnectMode ConnectMode,
          /*[in]*/ IDispatch * AddInInst,
          /*[in]*/ SAFEARRAY * * custom) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnDisconnection(
          /*[in]*/ enum AddInDesignerObjects::ext_DisconnectMode RemoveMode,
          /*[in]*/ SAFEARRAY * * custom) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnAddInsUpdate(
          /*[in]*/ SAFEARRAY * * custom) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnStartupComplete(
          /*[in]*/ SAFEARRAY * * custom) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnBeginShutdown(
          /*[in]*/ SAFEARRAY * * custom) override
        {
          return S_OK;
        }
       
        IRibbonExtensibilityPtr ribbon;
      };

      RegisterCom<ComAddinImpl> _registrar;
      bool _connected = false;

      auto& comAddinImpl()
      {
        return _registrar.server();
      }

    public:
      ComAddinCreator(const wchar_t* name, const wchar_t* description)
        : _registrar(
            new CComObject<ComAddinImpl>(),
            formatStr(L"%s.ComAddin", name ? name : L"xlOil").c_str())
      {
        if (!name)
          XLO_THROW("Com add-in name must be provided");
        auto addinPath = fmt::format(
          L"Software\\Microsoft\\Office\\Excel\\AddIns\\{0}", _registrar.progid());
        _registrar.writeRegistry(
          HKEY_CURRENT_USER, addinPath.c_str(), L"FriendlyName", name);
        _registrar.writeRegistry(
          HKEY_CURRENT_USER, addinPath.c_str(), L"LoadBehavior", (DWORD)0);
        if (description)
          _registrar.writeRegistry(
            HKEY_CURRENT_USER, addinPath.c_str(), L"Description", description);
      }

      void connect() override
      {
        if (_connected)
          return;
        try
        {
          auto& app = excelApp();
          // TODO: set this guy back!
          app.AutomationSecurity = Office::MsoAutomationSecurity::msoAutomationSecurityLow;
          app.GetCOMAddIns()->Update();
          auto ourProgid = _variant_t(progid());
          auto ourAddin = app.GetCOMAddIns()->Item(&ourProgid);
          ourAddin->Connect = VARIANT_TRUE;
          _connected = true;
        }
        XLO_RETHROW_COM_ERROR;
      }

      void disconnect() override
      {
        if (!_connected)
          return;
        try
        {
          auto& app = excelApp();
          auto ourProgid = _variant_t(progid());
          auto ourAddin = app.GetCOMAddIns()->Item(&ourProgid);
          ourAddin->Connect = VARIANT_FALSE;
          _connected = false;
        }
        XLO_RETHROW_COM_ERROR;
      }

      virtual void setRibbon(
        const wchar_t* xml,
        const std::map<std::wstring, std::function<void(const RibbonControl&)>> handlers)
      {
        if (_connected)
          XLO_THROW("Can only set Ribbon when add-in is disconnected");
        comAddinImpl().ribbon = createRibbon(xml, handlers);
      }

      ~ComAddinCreator()
      {
        try
        {
          disconnect();
        }
        catch (const std::exception& e)
        {
          XLO_THROW("ComAddin failed to close: {0}", e.what());
        }
      }

      const wchar_t* progid() const override
      {
        return _registrar.progid();
      }
    };

    std::shared_ptr<IComAddin> createComAddin(
      const wchar_t* name, const wchar_t* description)
    {
      return std::make_shared<ComAddinCreator>(name, description);
    }
  }
}