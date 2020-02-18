#include "XllContextInvoke.h"
#include "ExcelTypeLib.h"
#include "Connect.h"
#include "ExcelObj.h"
#include "CallHelper.h"
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Log.h>

namespace xloil
{

  static const std::function<void()>* theTargetFunc = nullptr;

  // TODO: make these commmands so they are hidden?
  XLO_ENTRY_POINT(XLOIL_XLOPER*) xloRunFuncInXLLContext()
  {
    // Do we need this result?
    static ExcelObj result;
    try
    {
      ScopeInXllContext context;
      (*theTargetFunc)();
    }
    catch (...)
    {
    }
    return &result;
  }
  XLO_REGISTER(xloRunFuncInXLLContext)
    .macro();

  static int theExcelCallFunc = 0;
  static XLOIL_XLOPER* theExcelCallResult = nullptr;
  static XLOIL_XLOPER** theExcelCallArgs = nullptr;
  static int theExcelCallNumArgs = 0;

  XLO_ENTRY_POINT(XLOIL_XLOPER*) xloRunInXLLContext()
  {
    static ExcelObj result(0);
    try
    {
      ScopeInXllContext context;
      result.val.w = Excel12v(theExcelCallFunc, theExcelCallResult, theExcelCallNumArgs, theExcelCallArgs);
    }
    catch (...)
    {
    }
    return &result;
  }
  XLO_REGISTER(xloRunInXLLContext)
    .macro();

  ScopeInXllContext::ScopeInXllContext()
  {
    ++_count;
  }
  ScopeInXllContext::~ScopeInXllContext()
  {
    --_count;
  }
  bool ScopeInXllContext::check()
  {
    return _count > 0;
  }

  int ScopeInXllContext::_count = 0;

  bool runInXllContext(const std::function<void()>& f)
  {
    if (ScopeInXllContext::check())
    {
      f();
      return true;
    }

    auto[result, xlret] = tryCallExcel(msxll::xlfGetDocument, 1);
    if (xlret == 0)
    {
      f();
      return true;
    }

    theTargetFunc = &f;

    auto ret = retryComCall([]() { return excelApp().Run("xloRunFuncInXLLContext"); });
    return ret.has_value();
  }

  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args)
  {
    if (ScopeInXllContext::check())
    {
      Excel12v(func, result, nArgs, (XLOIL_XLOPER**)args);
      return true;
    }
    theExcelCallFunc = func;
    theExcelCallResult = result;
    theExcelCallArgs = (XLOIL_XLOPER**)args;
    theExcelCallNumArgs = nArgs;
    //XLO_TRACE("Calling into XLL context fn= {0:#x}", (size_t)&fn);
    auto ret = retryComCall([]() { return excelApp().Run("xloRunInXLLContext"); });
    if (!ret)
      return msxll::xlretInvXlfn;
    auto variant = ret.value();
    
    if (SUCCEEDED(VariantChangeType(&variant, &variant, 0, VT_I4)))
      return variant.lVal;

    return msxll::xlretInvXlfn;
  }
}