#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xlOil/ExcelRange.h>
#include <xloil/StaticRegister.h>
#include <xlOil/Preprocessor.h>

using std::wstring;

namespace xloil
{
#define XLOCONCAT_NARGS 10
#define XLOCONCAT_ARG_NAME strings

  XLO_FUNC_START( xloConcat(
    const ExcelObj& separator, 
    XLO_DECLARE_ARGS(XLOCONCAT_NARGS, XLOCONCAT_ARG_NAME)
    )
  )
  {
    wstring result;
    wstring sep;
    auto pSeparator = nullptr;
     
    if (separator.isMissing())
    {
      ProcessArgs([&result](auto& argVal)
      {
        if (argVal.isNonEmpty())
          result += argVal.toString();
      }, XLO_ARGS_LIST(XLOCONCAT_NARGS, XLOCONCAT_ARG_NAME));
    }
    else
    {
      auto sep = separator.toString();
      ProcessArgs([&result, &sep](auto& argVal)
      {
        if (argVal.isNonEmpty())
          result += argVal.toString(sep.c_str()) + sep;
      }, XLO_ARGS_LIST(XLOCONCAT_NARGS, XLOCONCAT_ARG_NAME));
      if (!result.empty())
        result.erase(result.size() - sep.size());
    }
    return returnValue(result);
  }
  XLO_FUNC_END(xloConcat).threadsafe()
    .help(L"Concatenates strings. Non strings are converted to string, arrays are concatenated by row")
    .arg(L"Separator", L"[opt] separator between strings")
    XLO_WRITE_ARG_HELP(XLOCONCAT_NARGS, XLOCONCAT_ARG_NAME, L"Value or array of values to concatenate");
}