#include "ExcelObj.h"
#include "ArrayBuilder.h"
#include "ExcelArray.h"
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START(
    xloFillNA(
      const ExcelObj* value,
      ExcelObj* array,
      const ExcelObj* trim
    )
  )
  {
    if (!value->isType(ExcelType::ArrayValue))
      XLO_THROW("Value must be a suitable type for an array element");

    ExcelArray arr(*array, trim->isMissing() ? true : trim->toBool());

    const auto nRows = arr.nRows();
    const auto nCols = arr.nCols();

    // If the value type is string we need to copy the array, otherwise
    // we cheekily edit the input and pass it back to Excel. It seems 
    // that Excel copies the result before it frees in the input arguments.
    // Need to check that this behaviour is guaranteed....
    if (value->type() == ExcelType::Str)
    {
      auto valueStr = value->asPascalStr();

      // Rather than copy the string for each array entry, we just pass
      // the same pointer each time, so the total string length is 
      // just the length of the fill value.
      ExcelArrayBuilder builder(nRows, nCols, valueStr.length());

      // Set up the string data in the array memory space.
      auto arrayStr = builder.string(valueStr.length());
      arrayStr = valueStr;

      const auto arrayNumBytes = nRows * nCols * sizeof(ExcelObj);

      // Do a byte-wise copy of the entire array data - this should
      // be very fast. It means that string values will be pointing
      // to memory Excel allocated for the input argument, but when
      // we delete an array in xlOil, we don't recursively free over
      // each element so we won't double delete.
      auto newArray = builder.toExcelObj();
      memcpy_s(newArray.val.array.lparray, arrayNumBytes, 
               array->val.array.lparray, arrayNumBytes);

      // Now replace all N/As with the specified value. Note the builder
      // and our newArray both point to the same underlying data
      for (auto i = 0u; i < nRows; ++i)
        for (auto j = 0u; j < nCols; ++j)
          if (arr.at(i, j).isNA() || arr.at(i, j).isMissing())
            builder.emplace_at(i, j, PString<>(arrayStr.data()));

      return ExcelObj::returnValue(newArray);
    }   
    else
    {
      // For non strings, we simply replace N/As with the specified value.
      // ExcelArray is a view, so we const_cast to change the underlying data
      for (auto i = 0u; i < nRows; ++i)
        for (auto j = 0u; j < nCols; ++j)
          if (arr.at(i, j).isNA() || arr.at(i, j).isMissing())
            const_cast<ExcelObj&>(arr(i, j)) = *value;

      return array;
    }
  }
  XLO_FUNC_END(xloFillNA).threadsafe()
    .help(L"Replaces #N/As in the given array with a specifed value")
    .arg(L"Value", L"The value to fill with, must be a single type e.g. int, string, etc.")
    .arg(L"Array", L"The array containing #N/As")
    .arg(L"Trim", L"[opt] Specifies whether the array should be trimmed to the last row "
      "and column containing a non-#N/A and non-empty string value. Default is true");
}