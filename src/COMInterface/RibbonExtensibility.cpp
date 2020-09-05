#pragma once
#include "ExcelTypeLib.h"
#include "ClassFactory.h"
#include <xlOil/Log.h>
#include <xlOil/Interface.h>
#include <map>
#include <functional>
#include <regex>

using std::wstring;
using std::map;
using std::vector;

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

      vector<std::function<void(const RibbonControl&)>> _functions;
      map<wstring, DISPID> _idsOfNames;
      wstring _xml;
      CComPtr<IRibbonUI> _ribbonUI;

      // First two functions are raw_GetCustomUI and onLoadHandler
      static constexpr DISPID theFirstDispid = 3;

    public:
      RibbonImpl()
      {
      }

      virtual HRESULT __stdcall raw_GetCustomUI(
        /*[in]*/ BSTR RibbonID,
        /*[out,retval]*/ BSTR * RibbonXml) override
      {
        if (!_xml.empty())
          *RibbonXml = SysAllocString(_xml.data());
        return S_OK;
      }

      HRESULT onLoadHandler(IDispatch* ribbonUI)
      {
        IRibbonUI* ptr;
        if (ribbonUI->QueryInterface(&ptr) == S_OK)
          _ribbonUI.Attach(ptr);
        else
          XLO_ERROR("Ribbon load didn't work");
        return S_OK;
      }

      void setRibbon(
        const wchar_t* xml,
        const std::map<std::wstring, std::function<void(const RibbonControl&)>> handlers)
      {
        if (_ribbonUI)
          XLO_THROW("Already set"); // TODO: reload addin?
        std::wregex find(L"(<customUI[^>]*)>");
        _xml = std::regex_replace(xml, find, L"$1 onLoad=\"onLoadHandler\">");

        for (auto[name, fn] : handlers)
        {
          _functions.push_back(fn);
          _idsOfNames[name] = theFirstDispid - 1 + (DISPID)_functions.size();
        }
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

      STDMETHOD(GetTypeInfoCount)(_Out_ UINT* pctinfo)
      {
        return 0;
      }

      STDMETHOD(GetTypeInfo)(
        UINT itinfo,
        LCID lcid,
        _Outptr_result_maybenull_ ITypeInfo** pptinfo)
      {
        return E_NOTIMPL;
      }

      STDMETHOD(GetIDsOfNames)(
        _In_ REFIID riid,
        _In_reads_(cNames) _Deref_pre_z_ LPOLESTR* rgszNames,
        _In_range_(0, 16384) UINT cNames,
        LCID lcid,
        _Out_ DISPID* rgdispid)
      {
        if (cNames != 1)
          return DISP_E_UNKNOWNNAME;
        auto found = _idsOfNames.find(rgszNames[0]);
        if (found == _idsOfNames.end())
        {
          XLO_ERROR(L"Unknown handler '{0}' called by Ribbon", rgszNames[0]);
          return DISP_E_UNKNOWNNAME;
        }
        *rgdispid = found->second;
        return S_OK;
      }

      STDMETHOD(Invoke)(
        _In_ DISPID dispidMember,
        _In_ REFIID riid,
        _In_ LCID lcid,
        _In_ WORD wFlags,
        _In_ DISPPARAMS* pdispparams,
        _Out_opt_ VARIANT* pvarResult,
        _Out_opt_ EXCEPINFO* pexcepinfo,
        _Out_opt_ UINT* puArgErr)
      {
        auto* rgvarg = pdispparams->rgvarg;

        // TODO: handle pvarResult?
        if (dispidMember == 1)
        {
          return raw_GetCustomUI(rgvarg[1].bstrVal, rgvarg[0].pbstrVal);
        }
        else if (dispidMember == 2)
        {
          return onLoadHandler(rgvarg[1].pdispVal);
        }
        else if (dispidMember - theFirstDispid < _functions.size())
        {
          auto ctrl = (IRibbonControl*)rgvarg[0].pdispVal;
          try
          {
            (_functions[dispidMember - theFirstDispid])(RibbonControl{ ctrl->Id, ctrl->Tag });
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Error during ribbon callback: {0}", e.what());
            // set exception?
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

    Office::IRibbonExtensibility* createRibbon(
      const wchar_t* xml,
      const std::map<std::wstring, std::function<void(const RibbonControl&)>> handlers) 
    {
      auto ribbon = std::make_unique<CComObject<RibbonImpl>>();
      ribbon->setRibbon(xml, handlers);
      return ribbon.release();
    }
  }
}
