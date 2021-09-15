#pragma once
#include <xlOil/ExcelTypeLib.h>
#include "RibbonExtensibility.h"
#include "ClassFactory.h"
#include <xlOil/Log.h>
#include <xlOil/Ribbon.h>
#include <map>
#include <functional>
#include <regex>

using std::wstring;
using std::map;
using std::vector;
using std::shared_ptr;
using namespace Office;

namespace xloil
{
  namespace COM
  {
    class __declspec(novtable)
      RibbonImpl :
        public CComObjectRootEx<CComSingleThreadModel>,
        public NoIDispatchImpl<IRibbonExtensibility>
    {
    private:

      vector<RibbonCallback> _functions;
      map<wstring, DISPID> _idsOfNames;
      wstring _xml;
      std::function<RibbonCallback(const wchar_t*)> _handler;

      static constexpr DISPID theFirstDispid = 3;

    public:
      CComPtr<IRibbonUI> ribbonUI;

      RibbonImpl()
      {
        // Function DispIds are:
        //    1    raw_GetCustomUI
        //    2    onLoadHandler - a standard ribbonUI callback
        //    3... custom callbacks passed via the setRibbon/addCallback functions
        // 
        _idsOfNames[L"onLoadHandler"] = 2;
      }
      ~RibbonImpl()
      {}

      virtual HRESULT __stdcall raw_GetCustomUI(
        /*[in]*/ BSTR /*RibbonID*/,
        /*[out,retval]*/ BSTR * RibbonXml) override
      {
        if (!_xml.empty())
          *RibbonXml = SysAllocString(_xml.data());
        return S_OK;
      }

      HRESULT onLoadHandler(IDispatch* disp)
      {
        IRibbonUI* ptr;
        if (disp->QueryInterface(&ptr) == S_OK)
          ribbonUI.Attach(ptr);
        else
          XLO_ERROR("Ribbon load didn't work");
        return S_OK;
      }

      void setRibbon(
        const wchar_t* xml,
        const std::function<RibbonCallback(const wchar_t*)>& handler)
      {
        if (!_xml.empty())
          XLO_THROW("Already set"); // TODO: reload addin?
        std::wregex find(L"(<customUI[^>]*)>");
        _xml = std::regex_replace(xml, find, L"$1 onLoad=\"onLoadHandler\">");
        _handler = handler;
      }

      int addCallback(const wchar_t* name, RibbonCallback&& fn)
      {
        _functions.emplace_back(fn);
        auto dispid = theFirstDispid - 1 + (DISPID)_functions.size();
        _idsOfNames[name] = dispid;
        return dispid;
      }

      HRESULT _InternalQueryInterface(REFIID riid, void** ppv) throw()
      {
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == IID_IDispatch
          || riid == __uuidof(IRibbonExtensibility))
        {
          *ppv = this;
          AddRef();
          return S_OK;
        }
        return E_NOINTERFACE;
      }
#pragma region IDispatch

      STDMETHOD(GetIDsOfNames)(
        _In_ REFIID /*riid*/,
        _In_reads_(cNames) _Deref_pre_z_ LPOLESTR* rgszNames,
        _In_range_(0, 16384) UINT cNames,
        LCID /*lcid*/,
        _Out_ DISPID* rgdispid)
      {
        auto* fnName = rgszNames[0];
        if (cNames != 1)
          return DISP_E_UNKNOWNNAME;
        auto found = _idsOfNames.find(fnName);
        if (found == _idsOfNames.end())
        {
          try
          {
            auto func = _handler(fnName);
            if (func)
            {
              *rgdispid = addCallback(fnName, std::move(func));
              return S_OK;
            }
          }
          catch (const std::exception& e)
          {
            XLO_ERROR(L"Error finding handler '{0}': {1}", fnName, utf8ToUtf16(e.what()));
          }
          catch (...)
          {
          }
          XLO_ERROR(L"Unknown handler '{0}' called by Ribbon", fnName);
          return DISP_E_UNKNOWNNAME;
        }
        *rgdispid = found->second;
        return S_OK;
      }

      STDMETHOD(Invoke)(
        _In_ DISPID dispidMember,
        _In_ REFIID /*riid*/,
        _In_ LCID /*lcid*/,
        _In_ WORD /*wFlags*/,
        _In_ DISPPARAMS* pdispparams,
        _Out_opt_ VARIANT* pvarResult,
        _Out_opt_ EXCEPINFO* /*pexcepinfo*/,
        _Out_opt_ UINT* /*puArgErr*/)
      {
        // Remember the args are in reverse order
        auto* rgvarg = pdispparams->rgvarg;

        if (dispidMember == 1)
        {
          return raw_GetCustomUI(rgvarg[1].bstrVal, rgvarg[0].pbstrVal);
        }
        else if (dispidMember == 2)
        {
          return onLoadHandler(rgvarg[0].pdispVal);
        }
        else if (dispidMember - theFirstDispid < _functions.size())
        {
          const auto nArgs = pdispparams->cArgs;

          // Assign enough space: no Ribbon callback has this many args
          VARIANT* args[4];

          // First arg is the ribbon control
          auto ctrl = (IRibbonControl*)rgvarg[nArgs - 1].pdispVal;

          // Reverse order the other args
          for (auto i = 1u; i < nArgs; ++i)
            args[i - 1] = &rgvarg[nArgs - 1 - i];
          
          try
          {
            _functions[dispidMember - theFirstDispid](
              RibbonControl{ ctrl->Id, ctrl->Tag },
              pvarResult,
              nArgs - 1,
              args);
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Error during ribbon callback: {0}", e.what());
            // TODO: set exception?
          }
          catch (...)
          {
            return E_FAIL;
          }
        }
        else
        {
          XLO_ERROR("Internal Error: unknown dispid called on ribbon Invoke.");
          return E_FAIL;
        }
        return S_OK;
      }

#pragma endregion

    };

    class Ribbon : public IRibbon
    {
    public:
      Ribbon(
        const wchar_t* xml, 
        const std::function<RibbonCallback(const wchar_t*)>& handler)
      {
        _ribbon = new ComObject<RibbonImpl>();
        _ribbon->setRibbon(xml, handler);
      }
      void invalidate(const wchar_t* controlId) const override
      {
        if ((*_ribbon).ribbonUI)
        {
          if (controlId)
            (*_ribbon).ribbonUI->InvalidateControl(controlId);
          else
            (*_ribbon).ribbonUI->Invalidate();
        }
      }

      bool activateTab(const wchar_t* controlId) const override
      {
        return (*_ribbon).ribbonUI
          ? (*_ribbon).ribbonUI->ActivateTab(controlId)
          : false;
      }

      Office::IRibbonExtensibility* getRibbon() override
      {
        return _ribbon;
      }

      CComPtr<ComObject<RibbonImpl>> _ribbon;
    };
    shared_ptr<IRibbon> createRibbon(
      const wchar_t* xml,
      const std::function<RibbonCallback(const wchar_t*)>& handler)
    {
      return std::make_shared<Ribbon>(xml, handler);
    }
  }
}
