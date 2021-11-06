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
    auto fromRow = inFromRow.toInt(1);         
    auto fromCol = inFromCol.toInt(1);

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
      return returnValue(array(fromRow, fromCol));

    auto toRow = inToRow.toInt(); 
    auto toCol = inToCol.toInt(); 

    // Move to 1-based indexing
    if (toRow > 0) --toRow;
    if (toCol > 0) --toCol;

    const auto slice = array.slice(
      fromRow, fromCol, 
      toRow == 0 ? array.nRows() : toRow,
      toCol == 0 ? array.nCols() : toCol);

    // TODO: under certain circumstances the input array may
    // not need to be copied
    return returnValue(slice.toExcelObj());
  }
  XLO_FUNC_END(xloIndex).threadsafe()
    .help(L"Extends the INDEX function to xlOil refs and sub-arrays. Indices are 1-based. "
          L"Returns a single value if ToRow and ToCol are omitted otherwise returns an array.")
    .arg(L"ArrayOrRef", L"A range/array or and xlOil ref")
    .optArg(L"FromRow", L"Starting row, 1 if omitted. If negative counts back from last row")
    .optArg(L"FromCol", L"Starting column, 1 if omitted. If negative counts back from last column")
    .optArg(L"ToRow", L"End row, not inclusive. If omitted uses FromRow+1. If zero or negative counts back from last row")
    .optArg(L"ToCol", L"End column, not inclusive. If omitted uses FromCol+1. If zero or negative counts back from last column");
}