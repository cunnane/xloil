#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xloil/StaticRegister.h>
#include <boost/preprocessor/repeat_from_to.hpp>
#include <xlOil/Preprocessor.h>

using std::vector;
using std::tuple;

/*
Layout should be
"1, 2, 3 ; 4, 5 ; 6"
"1, 2, 3 ; 4, , 5"
TODO: support transpose with a t
*/
namespace xloil
{
#define XLOBLOCK_NARGS 30
#define XLOBLOCK_ARG_NAME arg
#define COPY_ARG(z,N,d) args[N - 1] = &arg##N;

  XLO_FUNC_START( xloBlock(
    const ExcelObj* layout,
    XLO_DECLARE_ARGS(XLOBLOCK_NARGS, XLOBLOCK_ARG_NAME)
  ))
  {
    constexpr size_t nArgs = XLOBLOCK_NARGS - 1;
    const ExcelObj* args[] = { XLO_ARG_PTRS(XLOBLOCK_NARGS, XLOBLOCK_ARG_NAME) };

    constexpr short NEW_ROW = -1;
    constexpr short PADDING = 0;

    const auto parse = layout->toString();

    if (parse.empty())
      XLO_THROW("No layout");

    using row_t = ExcelArray::row_t;
    using col_t = ExcelArray::col_t;

    auto* p = parse.data();

    row_t totalRows = 0, rowsForCurrent = 0;
    col_t totalCols = 0, colsForCurrent = 0;
    size_t totalStrLength = 0;

    vector<tuple<row_t, col_t, short>> spec;

    while (true)
    {
      wchar_t* nextP;
      auto argNum = (short)wcstol(p, &nextP, 10); 
      if (argNum > nArgs)
        XLO_THROW("Arg {i} out of range", argNum);
      p = nextP;

      if (argNum == 0)
        spec.emplace_back(0, 0, PADDING);
      else
      {
        row_t nRows = 1;
        col_t nCols = 1;
        auto* obj = args[argNum - 1];
        switch (obj->type())
        {
        case ExcelType::Multi:
        {
          obj->trimmedArraySize(nRows, nCols);
          auto arrData = obj->asArray();
          auto endArray = arrData + nRows * nCols;
          for (; arrData != endArray; ++arrData)
            totalStrLength += arrData->stringLength();
        }
        case ExcelType::Str:
          totalStrLength += obj->stringLength();
          break;
        default:
          break;
        }
        spec.emplace_back(nRows, nCols, argNum);
        rowsForCurrent = std::max(rowsForCurrent, nRows);
        colsForCurrent += nCols;
      }
      auto separator = wcspbrk(p, L",;");

      if (!separator || *separator == L';')
      {
        spec.emplace_back(rowsForCurrent, 0, NEW_ROW);
        totalCols = std::max(totalCols, colsForCurrent);
        totalRows += rowsForCurrent;
        colsForCurrent = 0;
        rowsForCurrent = 0;
      }

      if (!separator)
        break;

      p = separator + 1;
    }

    if (totalCols == 0 || totalRows == 0)
      return returnValue(CellError::NA);

    // TODO: avoid looping the spec twice!
    row_t row = 0;
    col_t col = 0;
    for (size_t i = 0; i < spec.size(); ++i)
    {
      auto[nRows, nCols, iArg] = spec[i];
      switch (iArg)
      {
      case NEW_ROW:
        col = 0;
        break;
      case PADDING:
      {
        auto j = i;
        // Walk through the specs until we find the next new row, adding
        // the required columns as we go
        while (std::get<2>(spec[++j]) != NEW_ROW)
          col += std::get<1>(spec[j]);

        // The padding shoud be the number of columns required for this row
        // less the number we found in the above loop
        spec[i] = std::make_tuple(0, totalCols - col, PADDING);
        i = j;
        col = 0;
        break;
      }
      default:
        col += nCols;
      }
    }

    ExcelArrayBuilder builder(totalRows, totalCols, totalStrLength);

    // TODO: is it possible to fill just the holes or is that too complicated?
    for (row_t i = 0; i < totalRows; ++i)
      for (col_t j = 0; j < totalCols; ++j)
        builder(i, j) = CellError::NA;

    row = 0; col = 0;
    for (auto[nRows, nCols, iArg] : spec)
    {
      switch (iArg)
      {
      case NEW_ROW:
        col = 0;
        row += nRows;
        break;
      case PADDING:
        col += nCols;
        break;
      default:
      {
        auto* obj = args[iArg - 1];
        switch (obj->type())
        {
        case ExcelType::Multi:
        {
          ExcelArray arr(*obj, nRows, nCols);

          for (row_t i = 0; i < nRows; ++i)
            for (col_t j = 0; j < nCols; ++j)
              builder(row + i, col + j) = arr(i, j);

          col += nCols;
          break;
        }
        default:
          builder(row, col) = *obj;
          ++col;
        }
      }
      }
    }

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloBlock).threadsafe()
    .help(L"Creates a block matrix given a layout specification. "
      "Layout has the form '1, 2, 3; 4, ,5' where numbers refer "
      "to the n-th argument (1-based). Semi-colon indicates a new row. "
      "Two consecutive commas mean a resized padded space. Whitespace is ignored.")
    .arg(L"Layout", L"String specifying how to layout the blocks");
}