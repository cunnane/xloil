#pragma once
#include "ExcelObj.h"
#include "xloil/Log.h"
namespace xloil
{
  class ExcelDict
  {
  public:
    ExcelDict(const ExcelObj& obj)
      : _base(obj)
    {
      int nCols;
      assert(obj.trimmedArraySize(_rows, nCols));
      if (nCols != 2)
        XLO_THROW("Expecting a 2 column array");
      // TODO: enforce key is int or str
    }
    // TODO: lookup?
    const ExcelObj& item(const int i) const
    {
      return data()[i * baseCols() + 1];
    }
    const ExcelObj& key(const int i) const
    {
      return data()[i * baseCols() + 0];
    }
  private:
    const ExcelObj& _base;
    int _rows;
    int baseCols() const { return  _base.val.array.columns; }
    const ExcelObj* data() const { return (ExcelObj*)_base.val.array.lparray; }
  };
}