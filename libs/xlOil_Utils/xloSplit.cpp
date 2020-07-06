#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xlOil/ExcelRange.h>
#include <xloil/StaticRegister.h>
#include <xlOil/Preprocessor.h>

using std::wstring;
using std::vector;

namespace xloil
{
  XLO_FUNC_START(xloSplit(
    const ExcelObj& string,
    const ExcelObj& separators,
    const ExcelObj& consecutiveAsOne
  ))
  {
    auto consecutive = consecutiveAsOne.toBool(true);
    auto pstr = string.asPascalStr().data();
    auto view = string.asPascalStr().view();
    auto length = pstr[0];
    auto sep = separators.toString();

    vector<wchar_t> found;
    
    wchar_t prev = 0, next;
    while ((next = view.find(sep.c_str(), prev)) != (wchar_t)wstring::npos)
    {
      *pstr = next - prev;
      found.push_back(prev);

      if (consecutive)
      {
        while (sep.find(view[next]) != wstring::npos)
          if (++next == view.size())
            break;
      }
      else
        ++next;
      pstr += (next - prev);
      prev = next;
    }

    // Add suffix
    *pstr = length - prev;
    found.push_back(prev);

    // Reset pstr
    pstr = string.asPascalStr().data();
    ExcelArrayBuilder builder((uint32_t)found.size(), 1, length);
    for (auto i = 0; i < found.size(); ++i)
    {
      builder(i) = PString(pstr + found[i]);
    }

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloSplit).threadsafe()
    .help(L"")
    .arg(L"string", L"string")
    .arg(L"separators", L"separator between strings")
    .arg(L"consecutiveAsOne");
}