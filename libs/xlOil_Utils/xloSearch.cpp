#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xlOil/Range.h>
#include <xloil/StaticRegister.h>
#include <xlOil/Preprocessor.h>
#include <xloil/ExcelObjCache.h>
#include "RegexHelpers.h"

using std::wstring;
using std::vector;
using std::wstring_view;

namespace xloil
{
  XLO_FUNC_START(xloSearch(
    const ExcelObj& string,
    const ExcelObj& searchRegex,
    const ExcelObj& replaceExpr,
    const ExcelObj& ignoreCase,
    const ExcelObj& grammar
  ))
  {
    const auto& input = cacheCheck(string);
    if (!input.isType(ExcelType::Str))
      return returnValue(CellError::NA);

    const auto noCase = ignoreCase.get<bool>(false);

    auto regexOptions = parseRegexGrammar(grammar.cast<PStringRef>());
    if (ignoreCase.get<bool>(false))
      regexOptions |= std::regex_constants::icase;

    const std::wregex expression(searchRegex.toString(), regexOptions);

    const auto giveIndices = replaceExpr.getIf<double>() == -1;
    const auto replaceExpression = giveIndices ? wstring() : replaceExpr.toString();
    const auto doReplace = !replaceExpression.empty();

    const auto nGroups = expression.mark_count();
    
    const auto pStr = input.cast<PStringRef>();

    vector<std::wcmatch> matchResults;
    size_t totalStrLength = 0;
    auto beginIterator = std::wcregex_iterator(pStr.begin(), pStr.end(), expression);
    auto endIterator = std::wcregex_iterator();
    for (auto i = beginIterator; i != endIterator; ++i)
    {
      matchResults.push_back(*i);
      totalStrLength += (size_t)i->length();
    }

    const auto N = (ExcelObj::row_t)matchResults.size();
    if (N == 0)
      return returnValue(CellError::NA);

    const auto width = doReplace ? 1 : (nGroups == 0 ? 1 : nGroups);
    ExcelArrayBuilder builder(N, width, giveIndices ? 0 : totalStrLength);

    if (doReplace)
    {
      for (ExcelObj::row_t i = 0; i < N; ++i)
        builder(i, 0) = matchResults[i].format(replaceExpression);
    }
    else if (nGroups > 0)
    {
      if (giveIndices)
      {
        for (ExcelObj::row_t i = 0; i < N; ++i)
          for (size_t j = 0; j < nGroups; ++j)
            builder(i, j) = matchResults[i][j + 1].first - pStr.begin();
      }
      else
      {
        for (ExcelObj::row_t i = 0; i < N; ++i)
          for (size_t j = 0; j < nGroups; ++j)
            builder(i, j) = strView(matchResults[i][j + 1]);
      }
    }
    else 
    {
      if (giveIndices)
      {
        for (ExcelObj::row_t i = 0; i < N; ++i)
          builder(i, 0) = strView(matchResults[i][0]);
      }
      else
      {
        for (ExcelObj::row_t i = 0; i < N; ++i)
          builder(i, 0) = matchResults[i][0].first - pStr.begin();
      }
    }

    return returnValue(builder.toExcelObj());
      
  }
  XLO_FUNC_END(xloSearch).threadsafe()
    .help(L"Matches the given regex to provided strings. Default syntax is ECMA")
    .arg(L"String", L"String or 1-d array of strings. The output orientation will match the input.")
    .arg(L"Search", L"Search regex. If no capture groups are defined, the function ouputs match success TRUE/FALSE")
    .optArg(L"Replace", L"Defines the output string using capture groups. If empty, one capture group is shown per column. If -1, indices are output")
    .optArg(L"IgnoreCase", L"(false) Determines if letters match either case")
    .optArg(L"Grammar", _GrammerHelp);
}