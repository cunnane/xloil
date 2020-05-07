#include "LocalFunctions.h"
#include <xlOil/StaticRegister.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelRange.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Log.h>
#include <COMInterface/WorkbookScopeFunctions.h>
#include <boost/preprocessor/repetition/repeat_from_to.hpp>
#include <boost/preprocessor/repetition/enum_params.hpp>
#include <map>
using std::wstring;
using std::map;
using std::shared_ptr;
using std::vector;

#define XLOIL_LOCAL_RANGE_FUNC xloil_local_range_200
namespace xloil
{
  struct Closure
  {
    ExcelFuncObject func;
    std::shared_ptr<const FuncInfo> info;
  };

  // ensure this is cleaned before things close.

  map < wstring, map<wstring, Closure>> theRegistry;

  Closure& findOrThrow(const wchar_t* wbName, const wchar_t* funcName)
  {
    auto wb = theRegistry.find(wbName);
    if (wb == theRegistry.end())
      XLO_THROW(L"Workbook '{0}' unknown for local function '{1}'", wbName, funcName);
    auto func = wb->second.find(funcName);
    if (func == wb->second.end())
      XLO_THROW(L"Local function '{1}' in workbook '{0}' not found", wbName, funcName);
    return func->second;
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
    const std::vector<std::shared_ptr<const FuncInfo>>& registeredFuncs,
    const std::vector<ExcelFuncObject> funcs)
  {
    auto& wbFuncs = theRegistry[workbookName];
    wbFuncs.clear();
    vector<wstring> coreRedirects(registeredFuncs.size());
    for (size_t i = 0; i < registeredFuncs.size(); ++i)
    {
      auto& info = registeredFuncs[i];
      if (info->numArgs() > 28)
        XLO_ERROR(L"Local function {0} has more than 28 arguments", info->name);
      wbFuncs[info->name] = Closure{ funcs[i], info };
      bool usesRanges = std::any_of(info->args.begin(), info->args.end(),
        [](auto& p) { return p.allowRange; });
      coreRedirects[i] = fmt::format(L"xloil_local_{1}{0}", info->numArgs(), usesRanges ? L"range_" : L"");
    }
    writeLocalFunctionsToVBA(workbookName, registeredFuncs, coreRedirects);
  }

  void forgetLocalFunctions(const wchar_t* workbookName)
  {
    theRegistry.erase(workbookName);
  }
}

using namespace xloil;

template<class... Args>
ExcelObj* doFunc(const ExcelObj& workbook, const ExcelObj& function, Args&&... args)
{
  try
  {
    if constexpr (sizeof...(Args) > 0)
      const ExcelObj* params[] = { args... };
    else
      const ExcelObj* params[] = { nullptr };
    const auto& closure = xloil::findOrThrow(workbook.toString().c_str(), function.toString().c_str());
    return closure.func(*closure.info, params);
  }
  catch (const std::exception& e)
  {
    return ExcelObj::returnValue(e.what());
  }
}

template<class... Args>
ExcelObj* doFuncRange(const ExcelObj& workbook, const ExcelObj& function, Args&&... args)
{
  try
  {
    const ExcelObj* params[] = { args... };
    const auto& closure = xloil::findOrThrow(workbook.toString().c_str(), function.toString().c_str());
    const auto& info = closure.info;
    const auto nArgs = info->numArgs();
    std::list<ExcelObj> rangeConversions;
    for (size_t i = 0; i < nArgs; ++i)
    {
      if (params[i]->isType(ExcelType::RangeRef) && !info->args[i].allowRange)
      {
        rangeConversions.emplace_back();
        callExcelRaw(msxll::xlCoerce, &rangeConversions.back(), params[i]);
        params[i] = &rangeConversions.back();
      }
    }
    return closure.func(*closure.info, params);
  }
  catch (const std::exception& e)
  {
    return ExcelObj::returnValue(e.what());
  }
}

XLO_ENTRY_POINT(XLOIL_XLOPER*) xloil_local_0(const ExcelObj& workbook, const ExcelObj& function)
{
  return doFunc(workbook, function);
}
XLO_REGISTER_FUNC(xloil_local_0).macro().hidden();

#define XLOIL_LOCAL(N, impl, name) \
  XLO_ENTRY_POINT(ExcelObj*) name##_##N( \
    const ExcelObj& workbook, const ExcelObj& function, \
    BOOST_PP_ENUM_PARAMS(N, const ExcelObj& arg) )\
  { \
    return impl(workbook, function, BOOST_PP_ENUM_PARAMS(N, &arg)); \
  } \
  XLO_REGISTER_FUNC(name##_##N).macro().hidden()

#define RPT(z, N, data) XLOIL_LOCAL(N, doFunc, xloil_local);
BOOST_PP_REPEAT_FROM_TO(1, 28, RPT, data)
#undef RPT

#define RPT(z, N, data) XLOIL_LOCAL(N, doFuncRange, xloil_local_range).allowRange();
BOOST_PP_REPEAT_FROM_TO(1, 28, RPT, data)
#undef RPT
