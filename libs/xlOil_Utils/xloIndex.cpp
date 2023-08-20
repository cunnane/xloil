#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelObjCache.h>

namespace xloil
{
  XLO_FUNC_START(
    xloIndex(
      const ExcelObj& inArrayOrRef,
      const ExcelObj& inFromRow,
      const ExcelObj& inFromCol,
      const ExcelObj& inToRow,
      const ExcelObj& inToCol)
  )
  {
    // TODO: handle range

    ExcelArray array(cacheCheck(inArrayOrRef));
    auto fromRow = inFromRow.get<int>(1);         
    auto fromCol = inFromCol.get<int>(1);

    if (fromRow > 0)
      --fromRow;
    else 
      fromRow += array.nRows();

    if (fromCol > 0)
      --fromCol;
    else 
      fromCol += array.nCols();

    // If only the first three arguments are supplied, behave like the INDEX function
    if (inToRow.isMissing() && inToCol.isMissing())
    {
      if (inFromCol.isMissing())
        return returnValue(array.slice(fromRow, 0, fromRow + 1, array.nCols()).toExcelObjUnsafe());
      else if (inFromRow.isMissing())
        return returnValue(array.slice(0, fromCol, array.nRows(), fromCol + 1).toExcelObjUnsafe());
      else
        return returnValue(array(fromRow, fromCol));
    }

    int toRow, toCol;

    if (inToRow.isMissing())
      toRow = fromRow + 1;
    else
    {
      toRow = inToRow.get<int>();
      if (toRow > 0) --toRow;
    }

    if (inToCol.isMissing())
      toCol = fromCol + 1;
    else
    {
      toCol = inToCol.get<int>();
      if (toCol > 0) --toCol;
    }

    const auto slice = array.slice(
      fromRow, fromCol, 
      toRow == 0 ? array.nRows() : toRow,
      toCol == 0 ? array.nCols() : toCol);

    return returnValue(slice.toExcelObjUnsafe());
  }
  XLO_FUNC_END(xloIndex).threadsafe()
    .help(L"Extends the INDEX function to xlOil refs and sub-arrays. Indices are 1-based. "
          L"Returns a single value if exactly FromRow and ToRow are given. Returns a"
          L"column or row array if one is given. Returns a subarray if ToRow or ToCol"
          L"are also given.")
    .arg(L"ArrayOrRef", L"A range/array or an xlOil ref")
    .optArg(L"FromRow", L"Starting row. If negative, counts back from last row. If ommitted, the entire column is returned")
    .optArg(L"FromCol", L"Starting column. If negative, counts back from last column. If ommitted, the entire row is returned")
    .optArg(L"ToRow", L"End row, not inclusive. If omitted uses FromRow+1 if FromRow was gven. "
                      L"Zero is the last row.If negative, counts back from last row")
    .optArg(L"ToCol", L"End column, not inclusive. If omitted uses FromCol+1 if FromCol was given. "
                      L"Zero is the last column. If negative, counts back from last column");
}