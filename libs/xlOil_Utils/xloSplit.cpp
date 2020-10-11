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
  namespace
  {
    void findSplitPoints(
      vector<wchar_t>& found, 
      PStringView<>& input,  // This string will be modified!
      const wstring& sep,
      const bool consecutiveAsOne)
    {
      const auto view = input.view();
      auto pstr = input.data();
      const auto length = pstr[0];

      size_t prev = 0, next;
      while ((next = view.find(sep.c_str(), prev)) != wstring::npos)
      {
        *pstr = (wchar_t)(next - prev);
        found.push_back((wchar_t)prev);

        if (consecutiveAsOne)
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
      *pstr = length - (wchar_t)prev;
      found.push_back((wchar_t)prev);
    }
  }

  XLO_FUNC_START(xloSplit(
    const ExcelObj& stringOrArray,
    const ExcelObj& separators,
    const ExcelObj& consecutiveAsOne
  ))
  {
    // Note this functon relies on the currently observed fact that 
    // Excel doesn't mind if we modify the input data a little bit,
    // then pass it back as the return value to avoid copies. That 
    // is, it appears to copy the function result before freeing the
    // memory associated with the inputs.
    auto consecutive = consecutiveAsOne.toBool(true);
    auto sep = separators.toString();

    if (stringOrArray.isType(ExcelType::Multi))
    {
      ExcelArray inputArray(stringOrArray);
      if (inputArray.dims() != 1)
        XLO_THROW("Input array must be 1-dim");

      // Location of the sub-string start points
      vector<vector<wchar_t>> found(inputArray.size());
      size_t totalStrLength = 0, maxTokens = 1;
      size_t iVal = 0;
      for (auto& val : inputArray)
      {
        if (val.isType(ExcelType::Str))
        {
          auto pStr = val.asPascalStr();
          totalStrLength += pStr.length();
          findSplitPoints(found[iVal], pStr, sep, consecutive);
          maxTokens = std::max(maxTokens, found[iVal].size());
        }
        ++iVal;
      }

      // Orient output array consistent with input
      bool byRow = inputArray.nCols() == 1;

      ExcelArrayBuilder builder(
        byRow ? inputArray.size() : maxTokens, 
        byRow ? maxTokens : inputArray.size(),
        totalStrLength);

      // We don't intend to write to every cell, so need to initialise
      builder.fillNA();

      if (byRow)
      {
        for (auto i = 0; i < found.size(); ++i)
        {
          if (found[i].empty()) // No tokens or was not a string
            builder(i, 0) = inputArray(i);
          else
          {
            // We are actually taking a pointer to part of the input string,
            // pretending we 'own' it, then emplacing the resulting ExcelObj
            // in the builder to avoid a copy. The emplacement uses move ctors
            // so the PString dtor will not be called on the 'owned' sub-string
            auto pStr = inputArray(i).asPascalStr();
            for (auto j = 0; j < found[i].size(); ++j)
              builder(i, j).emplace_pstr(pStr.data() + found[i][j]);
          }
        }
      }
      else
      {
        for (auto i = 0; i < found.size(); ++i)
        {
          if (found[i].empty()) // No tokens or was not a string
            builder(0, i) = inputArray(i);
          else
          {
            auto pStr = inputArray(i).asPascalStr();
            for (auto j = 0; j < found[i].size(); ++j)
              builder(j, i).emplace_pstr(pStr.data() + found[i][j]);
          }
        }
      }

      return returnValue(builder.toExcelObj());
    }
    else if (stringOrArray.isType(ExcelType::Str))
    {
      vector<wchar_t> found;

      auto pStr = stringOrArray.asPascalStr();
      findSplitPoints(found, pStr, sep, consecutive);

      ExcelArrayBuilder builder((uint32_t)found.size(), 1, pStr.length());
      for (auto i = 0; i < found.size(); ++i)
        builder(i).emplace_pstr(pStr.data() + found[i]);

      return returnValue(builder.toExcelObj());
    }
    else // Not a string or array, so do not modify
      return returnValue(stringOrArray);
  }
  XLO_FUNC_END(xloSplit).threadsafe()
    .help(L"Splits a string or array of strings on a given separator. The array must be"
           "1-dim.")
    .arg(L"String", L"String or array of strings. Any non-strings will be unmodified")
    .arg(L"Separators", L"Separators between strings: each character is interpreted "
                         "as a distinct separator")
    .optArg(L"ConsecutiveAsOne", L"(true) Treat consecutive delimiters as one");
}