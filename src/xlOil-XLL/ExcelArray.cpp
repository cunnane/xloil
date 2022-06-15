#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/Range.h>
#include <xloil/ArrayBuilder.h>

namespace xloil
{
  ExcelArray::ExcelArray(const ExcelObj& obj, bool trim)
  {
    if (obj.isType(ExcelType::Multi))
    {
      _data = (const ExcelObj*)obj.val.array.lparray;
      _baseCols = (col_t)obj.val.array.columns;
      if (trim)
        trimmedArraySize(obj, _rows, _columns);
      else
      {
        _rows = obj.val.array.rows;
        _columns = (col_t)obj.val.array.columns;
      }
    }
    else if (obj.isType(ExcelType::ArrayValue))
    {
      _data = &obj;
      _rows = 1;
      _columns = 1;
      _baseCols = 1;
    }
    else
      XLO_THROW(L"Type {0} not allowed as an array element", enumAsWCString(obj.type()));
  }

  ExcelObj ExcelArray::toExcelObj() const
  {
    // Single value return
    if (nCols() == 1 && nRows() == 1)
      return at(0);

    // Empty array
    if (dims() == 0)
      return ExcelObj();

    size_t strLen = 0;
    for (auto& v : (*this))
      strLen += v.stringLength();

    ExcelArrayBuilder builder(nRows(), nCols(), strLen);
    for (auto i = 0u; i < nRows(); ++i)
      for (auto j = 0u; j < nCols(); ++j)
        builder(i, j) = at(i, j);

    return builder.toExcelObj();
  }
  
  bool ExcelArray::trimmedArraySize(const ExcelObj& obj, row_t& nRows, col_t& nCols)
  {
    if ((obj.xtype() & msxll::xltypeMulti) == 0)
    {
      nRows = 0; nCols = 0;
      return false;
    }

    const auto& arr = obj.val.array;
    const auto start = (ExcelObj*)arr.lparray;
    nRows = arr.rows;
    nCols = arr.columns;

    auto p = start + nCols * nRows - 1;

    for (; nRows > 0; --nRows)
      for (int c = (int)nCols - 1; c >= 0; --c, --p)
        if (p->isNonEmpty())
          goto StartColSearch;

  StartColSearch:
    for (; nCols > 0; --nCols)
      for (p = start + nCols - 1; p < (start + nCols * nRows); p += arr.columns)
        if (p->isNonEmpty())
          goto SearchDone;

  SearchDone:
    return true;
  }
}