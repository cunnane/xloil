#include <xloil/ArrayBuilder.h>

namespace xloil
{
  detail::ArrayBuilderAlloc ExcelArrayBuilder::initialiseAllocator(
    row_t& nRows, col_t& nCols, size_t strLength, bool padTo2DimArray)
  {
    // Add the terminators and string counts to total length. Maybe 
    // not every cell will be a string so this is an over-estimate
    if (strLength > 0)
      strLength += nCols * nRows * 2;

    if (padTo2DimArray)
    {
      if (nRows == 1) nRows = 2;
      if (nCols == 1) nCols = 2;
    }

    auto arrSize = nRows * nCols;

    return detail::ArrayBuilderAlloc(arrSize, strLength);
  }

  ExcelArrayBuilder::ExcelArrayBuilder(
    row_t nRows, col_t nCols,
    size_t totalStrLength, bool padTo2DimArray)
    : _nRows(nRows)
    , _nColumns(nCols)
    , _allocator(initialiseAllocator(_nRows, _nColumns, totalStrLength, padTo2DimArray))
  {
    if (padTo2DimArray)
    {
      // Add padding
      if (nCols < _nColumns)
        for (row_t i = 0; i < nRows; ++i)
          (*this)(i, nCols) = CellError::NA;

      if (nRows < _nRows)
        for (col_t j = 0; j < _nColumns; ++j)
          (*this)(nRows, j) = CellError::NA;
    }
  }

  namespace
  {
    /// <summary>
    /// Repoint all strings to the correct place, give a contiguous array
    /// of pstring data
    /// </summary>
    void fixStrings(ExcelObj* pStart, size_t size, wchar_t* stringData)
    {
      const auto pEnd = pStart + size;
      auto pStr = detail::PStringStackIterator(stringData);
      for (; pStart != pEnd; ++pStart)
      {
        if (pStart->xltype == msxll::xltypeStr)
        {
          pStart->val.str.data = *pStr;
          ++pStr;
        }
      }
    }
  }

  ExcelObj SequentialArrayBuilder::toExcelObj()
  {
    if (_target != last())
      XLO_THROW("Array not fully populated during build");

    const auto bytesOfArray = _objects.size();
    const auto bytesOfStrings = sizeof(wchar_t) * _strings.size();
    auto arrayData = new char[bytesOfArray + bytesOfStrings];

    auto stringData = (wchar_t*)(arrayData + bytesOfArray);
    memcpy(arrayData, _objects.data(), bytesOfArray);
    wmemcpy(stringData, _strings.data(), _strings.size());

    fixStrings((ExcelObj*)arrayData, nRows() * nColumns(), stringData);

    return ExcelObj((ExcelObj*)arrayData, _nRows, _nColumns);
  }

  void SequentialArrayBuilder::copyToBuilder(
    detail::ArrayBuilderIterator targetBegin, detail::ArrayBuilderIterator targetEnd)
  {
    auto sourcePtr = (ExcelObj*)_objects.data();
    auto sourceEnd = last();
    auto pStr = detail::PStringStackIterator(_strings.data());

    for (; targetBegin != targetEnd && sourcePtr != sourceEnd; ++targetBegin, ++sourcePtr)
    {
      if (sourcePtr->xltype == msxll::xltypeStr)
      {
        auto pbuf = *pStr;
        targetBegin->copy_string(pbuf + 1, pbuf[0]);
        ++pStr;
      }
      else
        targetBegin->take(std::move(*sourcePtr));
    }
  }
}