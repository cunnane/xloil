#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelObj.h>
#include <cassert>
using namespace msxll;

namespace xloil
{
  const wchar_t* xlRetCodeToString(int ret, bool checkXllContext)
  {
    if (checkXllContext)
    {
      ExcelObj dummy;
      if (Excel12v(xlStack, &dummy, 0, nullptr) == xlretInvXlfn)
        return L"XLL function called outside XLL Context";
    }
    switch (ret)
    {
    case xlretSuccess:    return L"success";
    case xlretAbort:      return L"macro was stopped by the user ";
    case xlretInvXlfn:    return L"invalid function number, or calling function does not have permission to call the function or command";
    case xlretInvCount:   return L"invalid number of arguments";
    case xlretInvXloper:  return L"invalid XLOPER structure";
    case xlretStackOvfl:  return L"stack overflow";
    case xlretFailed:     return L"command failed";
    case xlretUncalced:   return L"attempt to read an uncalculated cell: this requires macro sheet permission";
    case xlretNotThreadSafe:  return L"not allowed during multi-threaded calc";
    case xlretInvAsynchronousContext: return L"invalid asynchronous function handle";
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
