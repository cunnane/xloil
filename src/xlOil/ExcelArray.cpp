#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelRange.h>
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
        obj.trimmedArraySize(_rows, _columns);
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

  ExcelObj ExcelArray::toExcelObj(const bool alwaysCopy) const
  {
    // Single value return
    if (nCols() == 1 && nRows() == 1)
      return at(0);

    // Empty array
    if (dims() == 0)
      return ExcelObj();

    // Point to source data if possible, avoid copy
    if (!alwaysCopy && _columns == _baseCols)
      return ExcelObj(_data, _rows, _columns);

    size_t strLen = 0;
    for (auto& v : (*this))
      strLen += v.stringLength();

    ExcelArrayBuilder builder(nRows(), nCols(), strLen);
    for (auto i = 0u; i < nRows(); ++i)
      for (auto j = 0u; j < nCols(); ++j)
        builder(i, j) = at(i, j);

    return builder.toExcelObj();
  }
}