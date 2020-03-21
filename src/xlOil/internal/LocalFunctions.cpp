#include "LocalFunctions.h"
#include <xlOil/StaticRegister.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelRange.h>
#include <COMInterface/WorkbookScopeFunctions.h>
#include <boost/preprocessor/repetition/repeat_from_to.hpp>
#include <boost/preprocessor/repetition/enum_params.hpp>
#include "ExcelObj.h"
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
    ExcelFuncPrototype func;
    std::shared_ptr<const FuncInfo> info;
  };

  // ensure this is cleaned before things close.

  map < wstring, map<wstring, Closure>> theRegistry;

  auto workbookCloseHandler = xloil::Event_WorkbookClose().bind(
    [](const wchar_t* wbName)
    {
      auto found = theRegistry.find(wbName);
      if (found != theRegistry.end())
        theRegistry.erase(found);
    });

  void registerLocalFuncs(
    const wchar_t* workbookName,
    const std::vector<std::shared_ptr<const FuncInfo>>& registeredFuncs,
    const std::vector<ExcelFuncPrototype> funcs)
  {
    auto& wbFuncs = theRegistry[workbookName];
    wbFuncs.clear();
    vector<wstring> coreRedirects(registeredFuncs.size());
    for (auto i = 0; i < registeredFuncs.size(); ++i)
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

using xloil::ExcelObj;

template<class... Args>
ExcelObj* doFunc(const ExcelObj& workbook, const ExcelObj& function, Args... args)
{
  try
  {
    if constexpr (sizeof...(Args) > 0)
      const ExcelObj* params[] = { args... };
    else
      const ExcelObj* params[] = { nullptr };
    auto& closure = xloil::theRegistry[workbook.toString()][function.toString()];
    return closure.func(*closure.info, params);
  }
  catch (const std::exception& e)
  {
    return ExcelObj::returnValue(e.what());
  }
}

template<class... Args>
ExcelObj* doFuncRange(const ExcelObj& workbook, const ExcelObj& function, Args... args)
{
  try
  {
    const ExcelObj* params[] = { args... };
    auto& closure = xloil::theRegistry[workbook.toString()][function.toString()];
    const auto& info = closure.info;
    const auto nArgs = info->numArgs();
    std::list<ExcelObj> rangeConversions;
    for (auto i = 0; i < nArgs; ++i)
    {
      if (params[i]->isRangeRef() && !info->args[i].allowRange)
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
XLO_REGISTER(xloil_local_0).macro();


#define XLOIL_LOCAL(N, impl, name) \
  XLO_ENTRY_POINT(ExcelObj*) name##_##N( \
    const ExcelObj& workbook, const ExcelObj& function, \
    BOOST_PP_ENUM_PARAMS(N, const ExcelObj& arg) )\
  { \
    return impl(workbook, function, BOOST_PP_ENUM_PARAMS(N, &arg)); \
  } \
  XLO_REGISTER(name##_##N).macro()

#define RPT(z, N, data) XLOIL_LOCAL(N, doFunc, xloil_local);
BOOST_PP_REPEAT_FROM_TO(1, 28, RPT, data)
#undef RPT

#define RPT(z, N, data) XLOIL_LOCAL(N, doFuncRange, xloil_local_range).allowRange();
BOOST_PP_REPEAT_FROM_TO(1, 28, RPT, data)
#undef RPT
