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
  // GLOBALS
  namespace
  {
    // Only write/read this from the main thread
    map<intptr_t, shared_ptr<const WorksheetFuncSpec>> theLocalFuncRegistry;

    using LocalFunctionMap = std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>>;

    void unregisterLocalFuncs(LocalFunctionMap& toRemove)
    {
      for (auto& [k, v] : toRemove)
        theLocalFuncRegistry.erase(v->registerId());
    }

    bool theIsExecutingLocalFunction = false;
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
    const wchar_t* vbaModuleName,
    const LocalFuncs action)
  {
    LocalFunctionMap toRemove;

    // WerRewrite the module if asked to or if the name of any new
    // function matches any existing one (because the parameters may
    // have changed)
    auto rewriteVBAModule = action != LocalFuncs::APPEND_MODULE;

    if (rewriteVBAModule)
      existing.swap(toRemove);
    
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

    const auto nNewFuncs = toRegister.size();

    if (rewriteVBAModule)
    {
      for (auto& f : existing)
        toRegister.push_back(f.second);
    }
    
    const auto iNewFuncsEnd = toRegister.begin() + nNewFuncs;
    for (auto i = toRegister.begin(); i != iNewFuncsEnd; ++i)
      existing.emplace(i->get()->info()->name, *i);

    runExcelThread([
        workbookName = wstring(workbookName), 
        vbaModuleName = wstring(vbaModuleName),
        funcsToWrite= move(toRegister),
        toRemove = move(toRemove), 
        rewriteVBAModule,
        action
    ]() mutable
    {
      if (action == LocalFuncs::CLEAR_MODULES)
        COM::removeExistingXlOilVBA(workbookName.c_str());

      unregisterLocalFuncs(toRemove);
      COM::writeLocalFunctionsToVBA(
        workbookName.c_str(), 
        funcsToWrite, 
        vbaModuleName.c_str(), 
        !rewriteVBAModule);
      for (auto& f : funcsToWrite)
        theLocalFuncRegistry.emplace(f->registerId(), f->spec());
    });

    // FuncInfo for Intellisense
    vector<shared_ptr<const FuncInfo>> funcInfos;
    for (auto& f : toRegister)
      funcInfos.emplace_back(f->info());

    // We send this a separate call because it requires the XLL API
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

  bool isExecutingLocalFunction() 
  {
    return theIsExecutingLocalFunction;
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

    auto found = theLocalFuncRegistry.find(*funcId);
    if (found == theLocalFuncRegistry.end())
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

    theIsExecutingLocalFunction = true;
    auto* result = func.call(xllArgPtr);
    theIsExecutingLocalFunction = false;

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
    theIsExecutingLocalFunction = false;
    *returnVal = COM::stringToVariant(e.what());
    return E_FAIL;
  }
}
