#include "LocalFunctions.h"
#include <xlOil/StaticRegister.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelRange.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Log.h>
#include <COMInterface/ComVariant.h>
#include <COMInterface/WorkbookScopeFunctions.h>
#include <boost/preprocessor/repetition/repeat_from_to.hpp>
#include <boost/preprocessor/repetition/enum_params.hpp>
#include <map>

using std::wstring;
using std::map;
using std::shared_ptr;
using std::vector;
using std::make_shared;

namespace xloil
{
  // ensure this is cleaned before things close.

  map < wstring, map<wstring, shared_ptr<const LambdaFuncSpec>>> theRegistry;

  const LambdaFuncSpec& findOrThrow(const wchar_t* wbName, const wchar_t* funcName)
  {
    auto wb = theRegistry.find(wbName);
    if (wb == theRegistry.end())
      XLO_THROW(L"Workbook '{0}' unknown for local function '{1}'", wbName, funcName);
    auto func = wb->second.find(funcName);
    if (func == wb->second.end())
      XLO_THROW(L"Local function '{1}' in workbook '{0}' not found", wbName, funcName);
    return *func->second;
  }

  auto workbookCloseHandler = Event::WorkbookAfterClose().bind(
    [](const wchar_t* wbName)
    {
      auto found = theRegistry.find(wbName);
      if (found != theRegistry.end())
        theRegistry.erase(found);
    });

  void registerLocalFuncs(
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const FuncInfo>>& funcInfo,
    const std::vector<ExcelFuncObject> funcs)
  {
    auto& wbFuncs = theRegistry[workbookName];
    wbFuncs.clear();
    vector<shared_ptr<const FuncSpec>> funcSpecs;
    for (size_t i = 0; i < funcInfo.size(); ++i)
    {
      auto& info = funcInfo[i];
      auto spec = make_shared<LambdaFuncSpec>(info, funcs[i]);
      funcSpecs.push_back(spec);
      wbFuncs[info->name] = spec;
    }
    COM::writeLocalFunctionsToVBA(workbookName, funcSpecs);
  }

  void forgetLocalFunctions(const wchar_t* workbookName)
  {
    theRegistry.erase(workbookName);
  }
}

using namespace xloil;
XLO_ENTRY_POINT(int) localFunctionEntryPoint(
  const VARIANT* workbookName,
  const VARIANT* funcName,
  VARIANT* returnVal,
  const VARIANT* args)
{
  try
  {
    VariantClear(returnVal);

    // TODO: check they really are strings!
    auto& func = findOrThrow(workbookName->bstrVal, funcName->bstrVal);

    const auto nArgs = func.info()->numArgs();
    vector<ExcelObj> xllArgs;
    vector<const ExcelObj*> argPtrs;
    xllArgs.reserve(nArgs);
    argPtrs.reserve(nArgs);

    if ((args->vt & VT_ARRAY) == 0)
    {
      XLO_THROW("Not implemnted yet");
    }
    else
    {
      auto pArray = args->parray;
      const auto dims = pArray->cDims;
      VARTYPE vartype;
      SafeArrayGetVartype(pArray, &vartype);
      if (vartype != VT_VARIANT || dims != 1)
        XLO_THROW("Expecting 1d array of variant for args");

      VARIANT* pData;
      if (FAILED(SafeArrayAccessData(pArray, (void**)&pData)))
        XLO_THROW("Failed accessing args array");
      
      std::shared_ptr<SAFEARRAY> arrayCloser(pArray, SafeArrayUnaccessData);

      const auto arrSize = pArray->rgsabound[0].cElements;
      if (arrSize != nArgs)
        XLO_THROW("Expecting {0} args, got {1}", nArgs, arrSize);

      for (auto i = 0u; i < arrSize; ++i)
      {
        xllArgs.emplace_back(COM::variantToExcelObj(pData[i], func.info()->args[i].allowRange));
        argPtrs.emplace_back(&xllArgs.back());
      }
    }

    auto result = func.call(&argPtrs[0]);

    *returnVal = COM::excelObjToVariant(*result);

    if ((result->xltype & msxll::xlbitDLLFree) != 0)
      delete result;

    return S_OK;
  }
  catch (const std::exception& e)
  {
    *returnVal = COM::stringToVariant(e.what());
    return E_FAIL;
  }
}
