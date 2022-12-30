#include "XllContextInvoke.h"
#include <xlOil/ExcelTypeLib.h>
#include "Connect.h"
#include <xlOil/ExcelObj.h>
#include <xlOil/AppObjects.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Log.h>

namespace xloil
{
  template <class TFunc>
  auto tryComCall(TFunc fn) -> typename std::invoke_result<TFunc>::type
  {
    try
    {
      return fn();
    }
    XLO_RETHROW_COM_ERROR;
  }

  static const std::function<bool()>* theBoolFunc = nullptr;
  static int theExcelCallFunc = 0;
  static XLOIL_XLOPER* theExcelCallResult = nullptr;
  static XLOIL_XLOPER** theExcelCallArgs = nullptr;
  static int theExcelCallNumArgs = 0;

  XLO_ENTRY_POINT(XLOIL_XLOPER*) xloRunInXLLContext()
  {
    static ExcelObj result(0);
    try
    {
      InXllContext context;
      if (theBoolFunc)
        result.val.w = (*theBoolFunc)() ? 1 : 0;
      else
        result.val.w = Excel12v(theExcelCallFunc, theExcelCallResult, theExcelCallNumArgs, theExcelCallArgs);
    }
    catch (...)
    {
    }
    return &result;
  }
  auto dummy = XLO_REGISTER_LATER(xloRunInXLLContext)
    .macro().hidden();

  InXllContext::InXllContext()
  {
    ++_count;
  }
  InXllContext::~InXllContext()
  {
    --_count;
  }
  bool InXllContext::check()
  {
    return _count > 0;
  }

  int InXllContext::_count = 0;

  bool runInXllContext(const std::function<bool()>& f)
  {
    // May go wrong in a multi-thread evironment.
    if (InXllContext::check())
    {
      return f();
    }

    // Crashes when called from window proc at startup - investigate?
    //auto[result, xlret] = tryCallExcel(msxll::xlfGetDocument, 1);
    //if (xlret == 0)

    theBoolFunc = &f;

    return tryComCall([]()
    {
      auto result = COM::attachedApplication().com().Run("xloRunInXLLContext");
      if (result.vt == VT_ERROR)
        XLO_THROW(L"COM Error {0:#x}", result.scode);
      return result;
    });
  }

  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args)
  {
    if (InXllContext::check())
    {
      return Excel12v(func, result, nArgs, (XLOIL_XLOPER**)args);
    }
    theBoolFunc = nullptr;
    theExcelCallFunc = func;
    theExcelCallResult = result;
    theExcelCallArgs = (XLOIL_XLOPER**)args;
    theExcelCallNumArgs = nArgs;

    auto ret = tryComCall([]()
    { 
      return COM::attachedApplication().com().Run("xloRunInXLLContext");
    });

    if (!ret)
      return msxll::xlretInvXlfn;
    
    if (SUCCEEDED(VariantChangeType(&ret, &ret, 0, VT_I4)))
      return ret.lVal;

    return msxll::xlretInvXlfn;
  }
}