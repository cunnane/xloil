#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelObj.h>
#include <ComInterface/XllContextInvoke.h>
#include <cassert>
using namespace msxll;

namespace xloil
{
  const wchar_t* xlRetCodeToString(int ret)
  {
    switch (ret)
    {
    case xlretSuccess: return L"success";
    case xlretAbort:    return L"macro halted";
    case xlretInvXlfn:    return L"invalid function number";
    case xlretInvCount:    return L"invalid number of arguments";
    case xlretInvXloper:    return L"invalid OPER structure";
    case xlretStackOvfl:   return L"stack overflow";
    case xlretFailed:   return L"command failed";
    case xlretUncalced:   return L"uncalced cell";
    case xlretNotThreadSafe:  return L"not allowed during multi-threaded calc";
    case xlretInvAsynchronousContext:  return L"invalid asynchronous function handle";
    case xlretNotClusterSafe:  return L"not supported on cluster";
    default:
      return L"unknown error";
    }
  }
  
  // TODO: currently unused, supposed to indicate functions which are safe to call outside XLL context
  // I'm not sure they are in fact safe!
  bool isSafeFunction(int funcNumber)
  {
    switch (funcNumber)
    {
    case xlFree:
    case xlStack:
    case xlSheetId:
    case xlSheetNm:
    case xlGetInst:
    case xlGetHwnd:
    case xlGetInstPtr:
    case xlAsyncReturn:
      return true;
    default:
      return false;
    }
  }

  XLOIL_EXPORT int callExcelRaw(
    int func, ExcelObj* result, size_t nArgs, const ExcelObj** args)
  {
    auto ret = Excel12v(func, result, (int)nArgs, (XLOIL_XLOPER**)args);
    if (result)
      result->fromExcel();
    return ret;
  }
}
