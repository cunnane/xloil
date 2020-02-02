#include "ExcelCall.h"
#include "ExcelObj.h"
#include "ComInterface/Connect.h"
#include <cassert>
using namespace msxll;

// TODO: implement
/*
string parseCallError(int xlret)
{
  if (xlret == xlretSuccess)
    return "";

  if (xlfn & xlCommand)
    debugPrintf("xlCommand | ");
  if (xlfn & xlSpecial)
    debugPrintf("xlSpecial | ");
  if (xlfn & xlIntl)
    debugPrintf("xlIntl | ");
  if (xlfn & xlPrompt)
    debugPrintf("xlPrompt | ");

  debugPrintf("%u) callback failed:", xlfn & 0x0FFF);

  // More than one error bit may be on

  if (xlret & xlretAbort)
  {
    debugPrintf(" Macro Halted\r");
  }

  if (xlret & xlretInvXlfn)
  {
    debugPrintf(" Invalid Function Number\r");
  }

  if (xlret & xlretInvCount)
  {
    debugPrintf(" Invalid Number of Arguments\r");
  }

  if (xlret & xlretInvXloper)
  {
    debugPrintf(" Invalid XLOPER12\r");
  }

  if (xlret & xlretStackOvfl)
  {
    debugPrintf(" Stack Overflow\r");
  }

  if (xlret & xlretFailed)
  {
    debugPrintf(" Command failed\r");
  }

  if (xlret & xlretUncalced)
  {
    debugPrintf(" Uncalced cell\r");
  }
  xlretNotThreadSafe
  An XLL worksheet function registered as thread safe attempted to call a C API function that is not thread safe.For example, a thread - safe function cannot call the XLM function xlfGetCell.

    256
    xlRetInvAsynchronousContext
    (Starting in Excel 2010) The asynchronous function handle is invalid.

    512
    xlretNotClusterSafe
    (Starting in Excel 2010) The call is not supported on clusters.
}
*/
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

  XLOIL_EXPORT int callExcelRaw(int func, ExcelObj* result, int nArgs, const ExcelObj** args)
  {
    auto ret = Excel12v(func, result, nArgs, (XLOIL_XLOPER**)args);
    // Likely cause of xlretInvXlfn is running outside XLL context
    if (ret == xlretInvXlfn)
    {
      ret = runInXllContext(func, result, nArgs, args);
    }
    result->fromExcel();
    return ret;
  }
}
