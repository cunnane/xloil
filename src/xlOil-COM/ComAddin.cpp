#include "ComAddin.h"
#include "ExcelTypeLib.h"
#include "ClassFactory.h"
#include "Connect.h"
#include "RibbonExtensibility.h"
#include <xlOil/State.h>
#include <xlOil/Log.h>
#include <xlOil/ApiCall.h>
#include <xlOil/Ribbon.h>
#include <map>
#include <functional>

using std::wstring;
using std::map;
using std::vector;
using std::shared_ptr;

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
          /*[in]*/ IDispatch * /*Application*/,
          /*[in]*/ enum AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/,
          /*[in]*/ IDispatch * /*AddInInst*/,
          /*[in]*/ SAFEARRAY * * /*custom*/) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnDisconnection(
          /*[in]*/ enum AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/,
          /*[in]*/ SAFEARRAY * * /*custom*/) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnAddInsUpdate(
          /*[in]*/ SAFEARRAY * * /*custom*/) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnStartupComplete(
          /*[in]*/ SAFEARRAY * * /*custom*/) override
        {
          return S_OK;
        }
        virtual HRESULT __stdcall raw_OnBeginShutdown(
          /*[in]*/ SAFEARRAY * * /*custom*/) override
        {
          return S_OK;
        }
       
        IRibbonExtensibility* ribbon;
      };

      auto& comAddinImpl()
      {
        return _registrar.server();
      }

      RegisterCom<ComAddinImpl> _registrar;
      bool _connected = false;
      shared_ptr<IRibbon> _ribbon;

    public:
      ComAddinCreator(const wchar_t* name, const wchar_t* description)
        : _registrar(
          new CComObject<ComAddinImpl>(),
          formatStr(L"%s.ComAddin", name ? name : L"xlOil").c_str())
      {
        if (!name)
          XLO_THROW("Com add-in name must be provided");

        // It's possible the addin has already been registered and loaded and 
        // is just being reinitialised.
        if (isConnected())
        {
          disconnect();
        }
        else
        {
          auto addinPath = fmt::format(
            L"Software\\Microsoft\\Office\\Excel\\AddIns\\{0}", _registrar.progid());
          _registrar.writeRegistry(
            HKEY_CURRENT_USER, addinPath.c_str(), L"FriendlyName", name);
          _registrar.writeRegistry(
            HKEY_CURRENT_USER, addinPath.c_str(), L"LoadBehavior", (DWORD)0);
          if (description)
            _registrar.writeRegistry(
              HKEY_CURRENT_USER, addinPath.c_str(), L"Description", description);
          // TODO: set this guy back!
          excelApp().AutomationSecurity = Office::MsoAutomationSecurity::msoAutomationSecurityLow;
          excelApp().GetCOMAddIns()->Update();
        }
      }
      COMAddIn* getAddin() const
      {
        auto& app = excelApp();
        // TODO: set this guy back!
        app.AutomationSecurity = Office::MsoAutomationSecurity::msoAutomationSecurityLow;
        auto ourProgid = _variant_t(progid());
        COMAddIn* ourAddin = 0;
        app.GetCOMAddIns()->raw_Item(&ourProgid, &ourAddin);
        return ourAddin;
      }
      bool isConnected() const
      {
        try
        {
          auto addin = getAddin();
          return addin
            ? (addin->Connect == VARIANT_TRUE)
            : false;
        }
        XLO_RETHROW_COM_ERROR;
      }
      void connect() override
      {
        if (_connected)
          return;
        try
        {
          auto addin = getAddin();
          if (!addin)
            XLO_THROW(L"Add-in connect: could not find addin '{0}'", progid());
          addin->Connect = VARIANT_TRUE;
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
          auto addin = getAddin();
          if (!addin)
            XLO_THROW(L"Add-in disconnect: could not find addin '{0}'", progid());
          addin->Connect = VARIANT_FALSE;
          _connected = false;
        }
        XLO_RETHROW_COM_ERROR;
      }

      virtual void setRibbon(
        const wchar_t* xml,
        const RibbonMap& mapper)
      {
        if (_connected)
          XLO_THROW("Can only set Ribbon when add-in is disconnected");
        _ribbon = createRibbon(xml, mapper);
        comAddinImpl().ribbon = _ribbon->getRibbon();
      }

      ~ComAddinCreator()
      {
        try
        {
          disconnect();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("ComAddin failed to close: {0}", e.what());
        }
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
    };

    std::shared_ptr<IComAddin> createComAddin(
      const wchar_t* name, const wchar_t* description)
    {
      return std::make_shared<ComAddinCreator>(name, description);
    }
  }
}