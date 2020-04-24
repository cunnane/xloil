#include "ExcelArray.h"
#include "ExcelObj.h"
#include <xlOil/ExcelRange.h>

namespace xloil
{
  ExcelArray::ExcelArray(const ExcelObj& obj, bool trim)
  {
    switch (obj.type())
    {
      case ExcelType::SRef:
      case ExcelType::Ref:
      case ExcelType::BigData:
      case ExcelType::Flow:
        XLO_THROW(L"Type {0} not allowed as an array element", enumAsWCString(obj.type()));

      case ExcelType::Multi:
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
        break;
      }
      default:
      {
        _data = &obj;
        _rows = 1;
        _columns = 1;
        _baseCols = 1;
        break;
      }
    }
  }
}