#pragma once
#include "ExcelObj.h"
#include <cassert>
#include "xloil/Log.h"
namespace xloil
{
  class ExcelArray
  {
  public:
    ExcelArray(const ExcelObj& obj, bool trim = true)
      : _base(obj)
    {
      if (!obj.trimmedArraySize(_rows, _columns))
        XLO_THROW("Expected an array");
    }

    const ExcelObj& operator()(int row, int col) const
    {
      return data()[row * baseCols() + col];
    }
    ExcelObj& operator()(int row, int col)
    {
      return data()[row * baseCols() + col];
    }
    const ExcelObj& operator()(int row) const
    {
      return data()[row];
    }
    ExcelObj& operator()(int row)
    {
      return data()[row];
    }
    int nRows() const { return _rows; }
    int nCols() const { return _columns; }
    size_t size() const { return _rows * _columns; }
    size_t dims() const { return _rows > 1 && _columns > 1 ? 2 : 1; }

    const ExcelObj* row_begin(int i) const { return data() + (i * baseCols()); }
    ExcelObj* row_begin(int i) { return data() + (i * baseCols()); }
    const ExcelObj* row_end(int i)   const { return data() + (i * baseCols() + nCols()); }
    ExcelObj* row_end(int i) { return data() + (i * baseCols() + nCols()); }

  private:
    const ExcelObj& _base;
    int _rows;
    int _columns;

    int baseCols() const { return _base.val.array.columns; }
    ExcelObj* data() { return (ExcelObj*)_base.val.array.lparray; }
    const ExcelObj* data() const { return (ExcelObj*)_base.val.array.lparray; }
  };

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
          emplace_at(i, nCols, CellError::NA);

      if (nRows < nPaddedRows)
        for (size_t j = 0; j < nPaddedCols; ++j)
          emplace_at(nRows, j, CellError::NA);
    }
    int emplace_at(size_t i, size_t j)
    {
      new (at(i, j)) ExcelObj();
      return 0;
    }
    // TODO: this is lazy, only int, bool, double and ExcelError are supported here, others are UB
    template <class T>
    int emplace_at(size_t i, size_t j, T&& x)
    {
      new (at(i, j)) ExcelObj(std::forward<T>(x));
      return 0;
    }
    int emplace_at(size_t i, size_t j, wchar_t*& buf, size_t& len)
    {
      buf = _stringData + 1;
      // TODO: check overflow?
      _stringData[0] = wchar_t(len);
      _stringData[len] = L'\0';
      new (at(i, j)) ExcelObj(PString<wchar_t>(_stringData));
      _stringData += len + 2;

      assert(_stringData <= _endStringData);
      return 0;
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

  private:
    ExcelObj * _arrayData;
    wchar_t* _stringData;
    const wchar_t* _endStringData;
    size_t _nRows, _nColumns;
  };
}