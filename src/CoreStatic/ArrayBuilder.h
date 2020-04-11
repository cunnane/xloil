#pragma once

#include "ExcelObj.h"
#include <xlOil/Throw.h>
#include <cassert>

namespace xloil
{
  class ExcelArrayBuilder
  {
  public:
    using row_t = uint32_t;
    using col_t = uint16_t;

    ExcelArrayBuilder(size_t nRows, size_t nCols,
      size_t totalStrLength = 0, bool pad2DimArray = false)
    {
      // Add the terminators and string counts to total length
      // Not everything has to be a string so this is an over-estimate
      if (totalStrLength > 0)
        totalStrLength += nCols * nRows * 2;

      auto nPaddedRows = (row_t)nRows;
      auto nPaddedCols = (col_t)nCols;
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
        for (row_t i = 0; i < nRows; ++i)
          setNA(i, nCols);

      if (nRows < nPaddedRows)
        for (col_t j = 0; j < nPaddedCols; ++j)
          setNA(nRows, j);
    }
    void setNA(size_t i, size_t j)
    {
      new (at(i, j)) ExcelObj(CellError::NA);
    }

    template<class T>
    void emplace_at(size_t i, size_t j, T&& x)
    {
      new (at(i, j)) ExcelObj(std::forward<T>(x));
    }

    void emplace_at(size_t i, size_t j, const ExcelObj& x)
    {
      auto type = x.type();
      if (((int)type & (int)ExcelType::ArrayValue) == 0)
        XLO_THROW("ExcelObj not of array value type");
      if (type == ExcelType::Str)
      {
        auto pstr = x.asPascalStr();
        emplace_at(i, j, pstr.begin(), pstr.length());
      }
      else
        new (at(i, j)) ExcelObj(x);
    }

    void emplace_at(size_t i, size_t j, wchar_t* str)
    {
      emplace_at(i, j, const_cast<const wchar_t*>(str));
    }

    void emplace_at(size_t i, size_t j, const wchar_t* str, int len = -1)
    {
      if (len < 0)
        len = (int)wcslen(str);
      auto pstr = string(len);
      wmemcpy_s(pstr.pstr(), len, str, len);
      auto xlObj = new (at(i, j)) ExcelObj();
      // We overwrite the object's string store directly, knowing its
      // d'tor will never be called as it is part of an array.
      xlObj->val.str = pstr.data();
      xlObj->xltype = msxll::xltypeStr;
    }

    PStringView<> string(size_t len)
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
      return PStringView<>(ptr);
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

    row_t nRows() const { return _nRows; }
    col_t nCols() const { return _nColumns; }

  private:
    ExcelObj* _arrayData;
    wchar_t* _stringData;
    const wchar_t* _endStringData;
    row_t _nRows;
    col_t _nColumns;
  };
}