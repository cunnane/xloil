#include "XllContextInvoke.h"
#include "ExcelTypeLib.h"
#include "Connect.h"
#include <xlOil/ExcelObj.h>
#include <xloil/ApiCall.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Log.h>

namespace xloil
{
  // See https://social.msdn.microsoft.com/Forums/vstudio/en-US/9168f9f2-e5bc-4535-8d7d-4e374ab8ff09/hresult-800ac472-from-set-operations-in-excel?forum=vsto
  constexpr HRESULT VBA_E_IGNORE = 0x800ac472;

  template <class TFunc>
  auto tryComCall(TFunc fn) -> typename std::invoke_result<TFunc>::type
  {
    try
    {
      return fn();
    }
    catch (_com_error& error)
    {
      if (error.Error() == VBA_E_IGNORE)
        throw ComBusyException("Excel COM is busy. A dialog box may be open. If this error persists, restart Excel.");
      else
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
    }
  }

  static const std::function<void()>* theVoidFunc = nullptr;
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
      if (theVoidFunc)
        (*theVoidFunc)();
      else
        result.val.w = Excel12v(theExcelCallFunc, theExcelCallResult, theExcelCallNumArgs, theExcelCallArgs);
    }
    catch (...)
    {
    }
    return &result;
  }
  XLO_REGISTER_FUNC(xloRunInXLLContext)
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
    return InComContext::_count > 0 ? false : _count > 0;
  }

  int InXllContext::_count = 0;

  InComContext::InComContext()
  {
    ++_count;
  }
  InComContext::~InComContext()
  {
    --_count;
  }
  bool InComContext::check()
  {
    return !InXllContext::check();
  }

  int InComContext::_count = 0;

  bool runInXllContext(const std::function<void()>& f)
  {
    // TODO: this whole Xll context thing may go wrong in a multi-thread evironment. Are we gteed to be on main?
    if (InXllContext::check())
    {
      f();
      return true;
    }

    // Crashes when called from window proc at startup - investigate?
    //auto[result, xlret] = tryCallExcel(msxll::xlfGetDocument, 1);
    //if (xlret == 0)

    theVoidFunc = &f;

    return tryComCall([]()
    {
      return COM::excelApp().Run("xloRunInXLLContext");
    });
  }

  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args)
  {
    if (InXllContext::check())
    {
      return Excel12v(func, result, nArgs, (XLOIL_XLOPER**)args);
    }
    theVoidFunc = nullptr;
    theExcelCallFunc = func;
    theExcelCallResult = result;
    theExcelCallArgs = (XLOIL_XLOPER**)args;
    theExcelCallNumArgs = nArgs;
    auto ret = tryComCall([]()
    { 
      return COM::excelApp().Run("xloRunInXLLContext");
    });
    if (!ret)
      return msxll::xlretInvXlfn;
    
    if (SUCCEEDED(VariantChangeType(&ret, &ret, 0, VT_I4)))
      return ret.lVal;

    return msxll::xlretInvXlfn;
  }
}