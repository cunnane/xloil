#include <xloil/ExcelObj.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <xloil/StaticRegister.h>

namespace xloil
{
  XLO_FUNC_START(
    xloPad(const ExcelObj* inArray)
  )
  {
    ExcelArray arr(*inArray);
    const auto nCols = arr.nCols();
    const auto nRows = arr.nRows();

    if (nCols > 1 && nRows > 1)
      return const_cast<ExcelObj*>(inArray);

    size_t strLen = 0u;
    for (auto& x : arr)
      strLen += x.stringLength();

    ExcelArrayBuilder builder(nRows, nCols, strLen, true);

    for (auto i = 0u; i < nRows; ++i)
      for (auto j = 0u; j < nCols; ++j)
        builder(i, j) = arr(i, j);

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloPad).threadsafe()
    .help(L"Returns at array with at least two rows and two columns, padded with #N/A")
    .arg(L"Array", L"The array or single value");
}