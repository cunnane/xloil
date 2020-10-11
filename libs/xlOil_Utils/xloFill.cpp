#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START( 
    xloFill(
      const ExcelObj* value, 
      const ExcelObj* nRows, 
      const ExcelObj* nCols)
  )
  {
    auto strLen = value->stringLength();
    auto nRowsVal = nRows->toInt();
    auto nColsVal = nCols->toInt();

    ExcelArrayBuilder builder(nRowsVal, nColsVal, strLen);

    if (value->type() == ExcelType::Str)
    {
      auto pstr = builder.string(strLen);
      pstr = value->asPascalStr();
      // Rather than copy the string for each array entry, we just pass
      // the same pointer each time.
      for (auto i = 0; i < nRowsVal; ++i)
        for (auto j = 0; j < nColsVal; ++j)
          builder(i, j).emplace_pstr(pstr.release());
    }
    else
    {
      for (auto i = 0; i < nRowsVal; ++i)
        for (auto j = 0; j < nColsVal; ++j)
          builder(i, j).copy_non_string(*value);
    }

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloFill).threadsafe()
    .help(L"Creates an array of the specified size filled with the given value")
    .arg(L"Value", L"The value to fill with, must be a single type e.g. int, string, etc.")
    .arg(L"NumRows")
    .arg(L"NumColumns");
}