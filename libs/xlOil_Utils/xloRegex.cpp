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
  namespace
  {
    template<class TStr>
    auto doMatch(const TStr& str, const std::wregex& regex, 
      std::wcmatch& results, bool isMatch)
    {
      return isMatch
        ? std::regex_match(
          str.begin(), str.end(),
          results, regex)
        : std::regex_search(
          str.begin(), str.end(),
          results, regex);
    }
  }

  XLO_FUNC_START(xloRegex(
    const ExcelObj& stringOrArray,
    const ExcelObj& searchRegex,
    const ExcelObj& replaceExpr,
    const ExcelObj& wholeString,
    const ExcelObj& ignoreCase,
    const ExcelObj& grammar
  ))
  {
    const auto isMatch = wholeString.get<bool>(false);

    auto regexOptions = parseRegexGrammar(grammar.cast<PStringRef>());
    if (ignoreCase.get<bool>(false))
      regexOptions |= std::regex_constants::icase;

    const std::wregex expression(searchRegex.toString(), regexOptions);

    const auto replaceExpression = replaceExpr.toString();
    const auto doReplace = !replaceExpression.empty();

    const auto nGroups = expression.mark_count();
    const auto& input = cacheCheck(stringOrArray);

    const auto notStringValue = nGroups == 0
      ? ExcelObj(false)
      : ExcelObj(CellError::NA);

    if (input.isType(ExcelType::Multi))
    {
      ExcelArray inputArray(input);
      if (inputArray.dims() != 1)
        XLO_THROW("Input array must be 1-dim");

      // Orient output array consistent with input
      const bool byRow = inputArray.nCols() == 1;

      // Do first parse to determine string length
      auto totalStrLength = 0u;
      for (auto& val : inputArray)
        if (val.isType(ExcelType::Str))
          totalStrLength += val.cast<PStringRef>().length();

      // Output array has 1 field if a replace string or no capture groups
      // have be specified, else one field per capture group
      const auto outputWidth = (doReplace || nGroups == 0) ? 1 : nGroups;

      ExcelArrayBuilder builder(
        byRow ? inputArray.size() : outputWidth,
        byRow ? outputWidth : inputArray.size(),
        totalStrLength);

      // We may not write to every cell, so need to initialise
      builder.fillNA();

      auto k = 0u;
      for (const auto& val : inputArray)
      {
        // First row/column entry
        auto writePosition = byRow ? builder(k, 0) : builder(0, k);

        if (val.isType(ExcelType::Str))
        {
          const auto pStr = val.cast<PStringRef>();

          std::wcmatch matchResults;
          const auto success = doMatch(pStr, expression, matchResults, isMatch);
          const auto N = matchResults.size();

          if (nGroups == 0)
            writePosition = success;
          else if (success)
          {
            assert(N > 1);

            if (doReplace)
              writePosition = matchResults.format(replaceExpression);
            else if (byRow)
              for (size_t j = 1; j < N; ++j)
                builder(k, j - 1) = strView(matchResults[j]);
            else
              for (size_t i = 1; i < N; ++i)
                builder(i - 1, k) = strView(matchResults[i]);
          }
        }
        else
        {
          writePosition = notStringValue;
        }
        ++k;
      }

      return returnValue(builder.toExcelObj());
    }
    else if (input.isType(ExcelType::Str))
    {
      const auto pStr = input.cast<PStringRef>();
      std::wcmatch matchResults;

      const auto success = doMatch(pStr, expression, matchResults, isMatch);

      const auto N = (ExcelObj::row_t)matchResults.size();

      if (nGroups == 0)
        return returnValue(success);
      else if (!success)
        return returnValue(CellError::NA);
      else if (doReplace)
        return returnValue(matchResults.format(replaceExpression));

      assert(N > 1);

      ExcelArrayBuilder builder(N - 1, 1, pStr.length());
      for (ExcelObj::row_t i = 1; i < N; ++i)
        builder(i - 1) = strView(matchResults[i]);

      return returnValue(builder.toExcelObj());
    }
    else
      return returnValue(notStringValue);
  }
  XLO_FUNC_END(xloRegex).threadsafe()
    .help(L"Matches the given regex against provided string(s). Default syntax is ECMA")
    .arg(L"String", L"String or 1-d array of strings. The output orientation will match the input.")
    .arg(L"Search", L"Search regex. If no capture groups are defined, the function ouputs match success: TRUE/FALSE")
    .optArg(L"Replace", L"If provided, defines the output string using capture groups, e.g '$1-$2'. If empty, one capture group is output per row/column")
    .optArg(L"WholeString", L"(false) If true, regex much match entire string, else string is searched for pattern")
    .optArg(L"IgnoreCase", L"(false) Determines if letters match either case")
    .optArg(L"Grammar", _GrammerHelp);
}