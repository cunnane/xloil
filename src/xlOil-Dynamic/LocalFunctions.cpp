#include "LocalFunctions.h"
#include <xlOil/StaticRegister.h>
#include <xlOil/DynamicRegister.h>
#include <xlOil/Range.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Log.h>
#include <xloil/FuncSpec.h>
#include <xlOil-COM/ComVariant.h>
#include <xlOil-COM/WorkbookScopeFunctions.h>
#include <xlOil-XLL/Intellisense.h>
#include <xloil/ExcelThread.h>
#include <oleacc.h>
#include <map>

using std::wstring;
using std::map;
using std::shared_ptr;
using std::vector;
using std::make_shared;
using std::move;

namespace xloil
{
  namespace
  {
    map<intptr_t, shared_ptr<const WorksheetFuncSpec>> theRegistry2;

    using LocalFunctionMap = std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>;

    void unregisterLocalFuncs(LocalFunctionMap& toRemove)
    {
      for (auto& [k, v] : toRemove)
        theRegistry2.erase(v->registerId());
    }
  }

  LocalWorksheetFunc::LocalWorksheetFunc(
    const std::shared_ptr<const WorksheetFuncSpec>& spec)
    : _spec(spec)
  {}

  LocalWorksheetFunc::~LocalWorksheetFunc()
  {
  }

  const std::shared_ptr<const WorksheetFuncSpec>& LocalWorksheetFunc::spec() const
  {
    return _spec;
  }
  const std::shared_ptr<const FuncInfo>& LocalWorksheetFunc::info() const
  {
    return _spec->info();
  }
  intptr_t LocalWorksheetFunc::registerId() const
  {
    return (intptr_t)_spec.get();
  }

  void registerLocalFuncs(
    LocalFunctionMap& existing,
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcs,
    const bool append)
  {
    LocalFunctionMap toRemove;

    if (!append)
      existing.swap(toRemove);

    auto rewriteVBAModule = !append;

    vector<shared_ptr<const LocalWorksheetFunc>> toRegister;

    for (auto& func : funcs)
    {
      toRegister.push_back(make_shared<LocalWorksheetFunc>(func));
      auto found = existing.find(func->name());
      if (found != existing.end())
      {
        rewriteVBAModule = true;
        toRemove.insert(existing.extract(found));
      }
    }

    auto iNewFuncsEnd = toRegister.end();

    if (rewriteVBAModule)
    {
      for (auto& f : existing)
        toRegister.push_back(f.second);
    }
    
    for (auto i = toRegister.begin(); i != iNewFuncsEnd; ++i)
      existing.emplace(i->get()->info()->name, *i);

    vector<shared_ptr<const FuncInfo>> funcInfos;
    for (auto& f : toRegister)
      funcInfos.emplace_back(f->info());

    runExcelThread([
        append, 
        workbookName = wstring(workbookName), 
        newRegisteredFuncs = move(toRegister),
        toRemove = move(toRemove) 
    ]() mutable
    {
      unregisterLocalFuncs(toRemove);
      COM::writeLocalFunctionsToVBA(workbookName.c_str(), newRegisteredFuncs, append);
      for (auto& f : newRegisteredFuncs)
        theRegistry2.emplace(f->registerId(), f->spec());
    });

    runExcelThread([funcInfos = std::move(funcInfos)]()
    {
      publishIntellisenseInfo(funcInfos);
    }, ExcelRunQueue::XLL_API | ExcelRunQueue::ENQUEUE);
  }

  void clearLocalFunctions(
    LocalFunctionMap& existing)
  {
    LocalFunctionMap toRemove;
    existing.swap(toRemove);
    runExcelThread([toRemove = move(toRemove)]() mutable
      {
        unregisterLocalFuncs(toRemove);
      });
  }
}

using namespace xloil;
int __stdcall localFunctionEntryPoint(
  const intptr_t* funcId,
  VARIANT* returnVal,
  const VARIANT* args)
{
  // This ensures the function is exported undecorated in x86 and x64
#pragma comment(linker, "/EXPORT:" __FUNCTION__"=" __FUNCDNAME__)

  try
  {
    VariantClear(returnVal);

   /* if (funcId->vt != VT_INT)
      XLO_THROW("WorkbookName and funcName parameters must be strings");*/

    auto found = theRegistry2.find(*funcId);
    if (found == theRegistry2.end())
      XLO_THROW("Local funcId {0} not found", *funcId);

    auto& func = *found->second;

    const auto nArgs = func.info()->numArgs();

    if ((args->vt & VT_ARRAY) == 0)
      XLO_THROW("Args must be an array");

    auto pArray = args->parray;
    const auto dims = pArray->cDims;
 
    if (dims != 1)
      XLO_THROW("Expecting 1d array of variant for 'args'");

    const auto arrSize = pArray->rgsabound[0].cElements;
    if (arrSize != nArgs)
      XLO_THROW("Expecting {0} args, got {1}", nArgs, arrSize);

    const ExcelObj** xllArgPtr = nullptr;
    vector<ExcelObj> xllArgs;
    vector<const ExcelObj*> argPtrs;

    if (arrSize > 0)
    {
      VARTYPE vartype;
      SafeArrayGetVartype(pArray, &vartype);
      if (vartype != VT_VARIANT)
        XLO_THROW("Expecting an array of variant for 'args'");

      VARIANT* pData;
      if (FAILED(SafeArrayAccessData(pArray, (void**)&pData)))
        XLO_THROW("Failed accessing 'args' array");

      std::shared_ptr<SAFEARRAY> arrayFinaliser(pArray, SafeArrayUnaccessData);

      xllArgs.reserve(nArgs);
      argPtrs.reserve(nArgs);

      for (auto i = 0u; i < arrSize; ++i)
      {
        xllArgs.emplace_back(
          COM::variantToExcelObj(pData[i], func.info()->args[i].type & FuncArg::Range));
        argPtrs.emplace_back(&xllArgs.back());
      }

      xllArgPtr = &argPtrs[0];
    }

    auto* result = func.call(xllArgPtr);

    // Commands (subroutines) can return a null pointer
    if (result)
    {
      COM::excelObjToVariant(returnVal, *result);

      if ((result->xltype & msxll::xlbitDLLFree) != 0)
        delete result;
    }

    return S_OK;
  }
  catch (const std::exception& e)
  {
    *returnVal = COM::stringToVariant(e.what());
    return E_FAIL;
  }
}
