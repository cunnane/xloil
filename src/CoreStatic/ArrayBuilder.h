#pragma once

#include "ExcelObj.h"
#include <cassert>

namespace xloil
{
  class ExcelArrayBuilder
  {
  public:
    ExcelArrayBuilder(size_t nRows, size_t nCols,
      size_t totalStrLength = 0, bool pad2DimArray = false)
    {
      // Add the terminators and string counts to total length
      // Not everything has to be a string so this is an over-estimate
      if (totalStrLength > 0)
        totalStrLength += nCols * nRows * 2;

      auto nPaddedRows = nRows;
      auto nPaddedCols = nCols;
      if (pad2DimArray)
      {
        if (nPaddedRows == 1) nPaddedRows = 2;
        if (nPaddedCols == 1) nPaddedCols = 2;
      }

      auto arrSize = nPaddedRows * nPaddedCols;

      auto* buf = new char[sizeof(ExcelObj) * arrSize + sizeof(wchar_t) * totalStrLength];
      _arrayData = (ExcelObj*)buf;
      _stringData = (wchar_t*)(_arrayData + arrSize);
      _endStringData = _stringData + totalStrLength;
      _nRows = nPaddedRows;
      _nColumns = nPaddedCols;

      // Add padding
      if (nCols < nPaddedCols)
        for (size_t i = 0; i < nRows; ++i)
          setNA(i, nCols);

      if (nRows < nPaddedRows)
        for (size_t j = 0; j < nPaddedCols; ++j)
          setNA(nRows, j);
    }
    void setNA(size_t i, size_t j)
    {
      new (at(i, j)) ExcelObj(CellError::NA);
    }
  
    template<class T>
    void emplace_at(size_t i, size_t j, T x) 
    { 
      new (at(i, j)) ExcelObj(x);
    }
    template<>
    void emplace_at<wchar_t*>(size_t i, size_t j, wchar_t* str)
    {
      emplace_at(i, j, const_cast<const wchar_t*>(str));
    }
    template<>
    void emplace_at<const wchar_t*>(size_t i, size_t j, const wchar_t* str)
    {
      auto len = wcslen(str);
      auto pstr = string(len);
      wmemcpy_s(pstr.pstr(), len, str, len);
      new (at(i, j)) ExcelObj(std::forward<PString<>>(pstr));
    }

    void emplace_at(size_t i, size_t j, ExcelObj&& x)
    {
      new (at(i, j)) ExcelObj(std::forward<ExcelObj>(x));
    }

    //void emplace_at(size_t i, size_t j, PString<>&& x)
    //{
    //  new (at(i, j)) ExcelObj(std::forward<PString<>>(x));
    //}

    PString<> string(size_t& len)
    {
      wchar_t* ptr = nullptr;
      if (len > 0)
      {
        assert(_stringData <= _endStringData);
        _stringData[0] = wchar_t(len);
        _stringData[len] = L'\0';
        ptr = _stringData;
        _stringData += len + 2;
      }
      return PString<wchar_t>::view(ptr);
    }

    ExcelObj* at(size_t i, size_t j)
    {
      assert(i < _nRows && j < _nColumns);
      return _arrayData + (i * _nColumns + j);
    }

    ExcelObj toExcelObj()
    {
      return ExcelObj(_arrayData, int(_nRows), int(_nColumns));
    }

    size_t nRows() const { return _nRows; }
    size_t nCols() const { return _nColumns; }

  private:
    ExcelObj * _arrayData;
    wchar_t* _stringData;
    const wchar_t* _endStringData;
    size_t _nRows, _nColumns;
  };
}