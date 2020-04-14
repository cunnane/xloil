#include "ExcelObj.h"
#include "ArrayBuilder.h"
#include "ExcelArray.h"
#include <xloil/StaticRegister.h>
#include <xlOil/Preprocessor.h>
#include <algorithm>
#include <numeric>

using std::array;
using std::vector;

namespace xloil
{
#define XLOSORT_NARGS 4
#define XLOSORT_ARG_NAME rowOrHeading
  namespace
  {
    enum SortDirection
    {
      Descending    = 1 << 0,
      CaseSensitive = 1 << 1,
      StopSearch    = 1 << 2
    };

    using MyArray = array<size_t, XLOSORT_NARGS + 1>;

    // TOOD: template<int N>
    struct LessThan
    {
      LessThan(const ExcelArray& data, const MyArray& directions, const MyArray& columns)
        : _data(data)
        , _directions(directions)
        , _columns(columns)
      {}
      bool operator()(const size_t left, const size_t right)
      {
        size_t i = 0;
        while (_directions[i] != StopSearch)
        {
          bool cased = _directions[i] & CaseSensitive;
          auto cmp = ExcelObj::compare(
            _data(left, _columns[i]), 
            _data(right, _columns[i]), 
            true, cased);
          if (cmp != 0)
            return (_directions[i] & Descending) == 0 ? cmp < 0 : cmp > 0;
          ++i;
        }
        return false;
      }
      const MyArray& _directions;
      const MyArray& _columns;
      const ExcelArray& _data;
    };

    void swapmem(size_t* a, size_t* b, size_t nBytes)
    {
      const auto end = (size_t*)((char*)a + nBytes);

      for (; a < end; ++a, ++b)
      {
        const auto t = *a;
        *a = *b;
        *b = t;
      }
    }
  }

  XLO_FUNC_START(
    xloSort(
      ExcelObj* array,
      const ExcelObj* order,
      const ExcelObj* rowOrHeading1,
      const ExcelObj* rowOrHeading2,
      const ExcelObj* rowOrHeading3,
      const ExcelObj* rowOrHeading4
    )
  )
  {
    ExcelArray arr(*array);
    const auto nRows = arr.nRows();
    const auto nCols = arr.nCols();

    const ExcelObj* args[] = { BOOST_PP_ENUM_SHIFTED_PARAMS(BOOST_PP_ADD(XLOSORT_NARGS,1), XLOSORT_ARG_NAME) };

    // could use raw pascal str, but unnecessary optimisation
    auto orderStr = order->toString(); 

    // Anything to do?
    if (orderStr.empty() || nRows < 2 || nCols == 0)
      return array;

    MyArray directions, columns;
    auto c = orderStr.begin();
    bool hasHeadings = false;
    auto i = 0;
    for (; i < directions.size() - 1; ++i, ++c)
    {
      while (c != orderStr.end() && iswspace(*c)) ++c;
      if (c == orderStr.end())
        break;

      directions[i] = 0;
      switch (*c)
      {
      case L'A':
        directions[i] |= CaseSensitive;
      case L'a':
        break;
      case L'D':
        directions[i] |= CaseSensitive;
      case L'd':
        directions[i] |= Descending;
        break;
      default:
        XLO_THROW("Direction must be one of {A, a, D, d}");
      }



      switch (args[i]->type())
      {
      case ExcelType::Int:
      case ExcelType::Num:
        columns[i] = args[i]->toInt() - 1; // 1-based column indexing to match Excel's INDEX function etc.
        if (columns[i] >= nCols)
          XLO_THROW("Column number in descriptor {0} is beyond number of array columns: {1} > {2}", 
            i, columns[i] + 1, nCols);
        break;
      case ExcelType::Str:
        hasHeadings = true;
        columns[i] = nCols;
        for (auto j = 0; j < nCols; ++j)
          if (*args[i] == arr(0, j))
          {
            columns[i] = j; 
            break;
          }
        if (columns[i] == nCols)
          XLO_THROW(L"Could not find heading {0} in first row of array", args[i]->toString());
        break;
      default:
        // No need to specify descriptor for a single column
        if (nCols == 1)
          break;
        XLO_THROW("Column descriptor {0} must be a column number or heading", i);
      }
    }
    directions[i] = StopSearch;

    using row_t = ExcelArray::row_t;

    vector<row_t> indices(nRows);
    std::iota(indices.begin(), indices.end(), 0);

    std::sort(indices.begin() + (hasHeadings ? 1 : 0), indices.end(),
      LessThan(arr, directions, columns));

    vector<row_t> cycle;
    const auto npos = row_t(-1);

    auto index = 0;
    while (true)
    {
      while (index < indices.size() && indices[index] == npos) ++index;
      if (index == indices.size())
        break;

      auto k = index;
      do
      {
        cycle.push_back(k);
        const auto r = k;
        k = indices[k];
        indices[r] = npos;
      } while (k != index);

      for (size_t i = 0; i < cycle.size() - 1; ++i)
      {
        swapmem(
          (size_t*)arr.row_begin(cycle[i]),
          (size_t*)arr.row_begin(cycle[i + 1]),
          nCols * sizeof(ExcelObj));
      }
      cycle.clear();
    }

    return array;
  }
  XLO_FUNC_END(xloSort).threadsafe()
    .help(L"")
    .arg(L"Array", L"")
    .arg(L"Order", L"")
    .arg(L"rowOrHeading1", L"[opt]");
}