#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xloil/StaticRegister.h>
#include <xlOil/Preprocessor.h>
#include <algorithm>
#include <numeric>

using std::array;
using std::vector;

namespace xloil
{
#define XLOSORT_NARGS 8
#define XLOSORT_ARG_NAME colOrHeading
  namespace
  {
    enum SortDirection
    {
      Descending    = 1 << 0,
      CaseSensitive = 1 << 1,
      StopSearch    = 1 << 2
    };

    using MyArray = array<ExcelArray::col_t, XLOSORT_NARGS + 1>;

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
            _data.at(left, _columns[i]), 
            _data.at(right, _columns[i]), 
            cased);
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
      XLO_DECLARE_ARGS(XLOSORT_NARGS, XLOSORT_ARG_NAME)
    )
  )
  {
    ExcelArray arr(*array);
    const auto nRows = arr.nRows();
    const auto nCols = arr.nCols();

    const ExcelObj* args[] = { XLO_ARG_PTRS(XLOSORT_NARGS, XLOSORT_ARG_NAME) };

    // could use raw pascal str, but unnecessary optimisation
    auto orderStr = order->toString(); 

    // Anything to do?
    if (orderStr.empty() || nRows < 2 || nCols == 0)
      return array;

    MyArray directions, columns;

    // Default sort order is left to right on columns
    std::iota(columns.begin(), columns.end(), 0);

    auto c = orderStr.begin();
    bool hasHeadings = false;
    auto nOrders = 0;
    for (; nOrders < directions.size() - 1; ++nOrders, ++c)
    {
      while (c != orderStr.end() && iswspace(*c)) ++c;
      if (c == orderStr.end())
        break;

      auto& direction = directions[nOrders];
      direction = 0;
      switch (*c)
      {
      case L'A':
        direction |= CaseSensitive;
      case L'a':
        break;
      case L'D':
        direction |= CaseSensitive;
      case L'd':
        direction |= Descending;
        break;
      default:
        XLO_THROW("Direction must be one of {A, a, D, d}");
      }

      const auto* arg = args[nOrders];
      auto& column = columns[nOrders];

      switch (arg->type())
      {
      case ExcelType::Int:
      case ExcelType::Num:
        column = arg->toInt() - 1; // 1-based column indexing to match Excel's INDEX function etc.
        if (column >= nCols)
          XLO_THROW("Column number in descriptor {0} is beyond number of array columns: {1} > {2}", 
            nOrders, column + 1, nCols);
        break;
      case ExcelType::Str:
        hasHeadings = true;
        column = nCols;
        for (auto j = 0; j < nCols; ++j)
          if (*arg == arr(0, j))
          {
            column = j;
            break;
          }
        if (column == nCols)
          XLO_THROW(L"Could not find heading {0} in first row of array", arg->toString());
        break;
      case ExcelType::Missing:
        // No need to specify descriptor: can rely on default ordering
        break;
      default:
        XLO_THROW("Column descriptor {0} must be a column number or heading", nOrders);
      }
    }
    directions[nOrders] = StopSearch;

    using row_t = ExcelArray::row_t;

    vector<row_t> indices(nRows);
    std::iota(indices.begin(), indices.end(), 0);

    std::sort(indices.begin() + (hasHeadings ? 1 : 0), indices.end(),
      LessThan(arr, directions, columns));

    // For an inplace sort, we note the indices array contains
    // the inverse of the permutation we need to apply to the rows
    // so we just step through each cycle, applying transpositions.
    // We mark moved rows with npos

    const auto npos = row_t(-1);

    row_t start = 0;
    while (true)
    {
      while (start < indices.size() && indices[start] == npos) ++start;
      if (start == indices.size())
        break;

      row_t k = start;
      while (true)
      {
        const auto r = indices[k];
        indices[k] = npos;
        if (r == start)
          break;
        swapmem(
          (size_t*)arr.row_begin(k),
          (size_t*)arr.row_begin(r),
          nCols * sizeof(ExcelObj));
        k = r;
      }
    }
    return array;
  }
  XLO_FUNC_END(xloSort).threadsafe()
    .help(L"Sorts an array by one or more columns. If column headings are specified the first row is "
      "not moved. The `Order` should contain one character for each column specified for sorting")
    .arg(L"Array", L"")
    .arg(L"Order", L"a = ascending, A = ascending case-sensitive, d = descending, D = descending "
      "case-sensitive, whitespace ignored")
    XLO_WRITE_ARG_HELP(XLOSORT_NARGS, XLOSORT_ARG_NAME, L"Column number (1-based) or column heading");
}