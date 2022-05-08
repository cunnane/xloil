
#include <xloil/StaticRegister.h>
#include <xloil-xll/FuncRegistry.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/Throw.h>
namespace xloil
{
  XLO_FUNC_START(xloHelp(
    const ExcelObj& function
  ))
  {
    const auto funcName = function.get<std::wstring_view>();
    // Why does wstring_view not work with find?
    const auto found = registeredFuncsByName().find(std::wstring(funcName));
    if (found == registeredFuncsByName().end())
      XLO_THROW(L"Function '{0}' not found", funcName);
    const auto func = found->second;
    const auto info = func->info();

    size_t stringLen = info->name.size() + info->help.size();
    const auto& args = info->args;
    const auto nArgs = args.size();
    for (auto& arg : args)
      stringLen += arg.name.size() + arg.help.size();
    
    ExcelArrayBuilder builder(1 + (ExcelObj::row_t)nArgs, 2, stringLen);
    builder(0, 0) = info->name;
    builder(0, 1) = info->help;
    for (auto i = 0u; i < nArgs; ++i)
    {
      builder(1 + i, 0) = args[i].name;
      builder(1 + i, 1) = args[i].help;
    }

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloHelp).threadsafe()
    .help(L"Give help on an function as 2-column array. "
      "First row is the function name and help string. Subsequent rows are "
      "argument names and their descriptions")
    .arg(L"function", L"Name of xloil registered function");
}