#pragma once
#include "ExcelObj.h"
#include <cassert>
#include <xloil/Log.h>
namespace xloil
{
  class ExcelArray
  {
  public:
    ExcelArray(const ExcelObj& obj, bool trim = true)
      : _data((ExcelObj*)obj.val.array.lparray)
      , _colOffset(0)
      , _baseCols(obj.val.array.columns)
    {
      if (obj.type() != ExcelType::Multi)
        XLO_THROW("Expected an array");

      if (trim)
        obj.trimmedArraySize(_rows, _columns);
      else
      {
        _rows = obj.val.array.rows;
        _columns = obj.val.array.columns;
      }
    }

    ExcelArray(const ExcelArray& arr, int fromRow, int fromCol, int toRow, int toCol)
      : _rows((toRow < 0 ? arr._rows + toRow : toRow) - fromRow)
      , _columns((toCol < 0 ? arr._columns + toCol : toCol) - fromCol)
      , _colOffset(fromCol)
      , _baseCols(arr._baseCols)
    {
      _data = arr._data + fromRow * _baseCols;
    }

    const ExcelObj& operator()(size_t row, size_t col) const
    {
      checkRange(row, col);
      return at(row, col);
    }
    ExcelObj& operator()(size_t row, size_t col)
    {
      checkRange(row, col);
      return at(row, col);
    }
    const ExcelObj& operator()(size_t row) const
    {
      //checkRange(row, 0);
      return at(row);
    }
    ExcelObj& operator()(size_t row)
    {
      //checkRange(row, 0);
      return at(row);
    }

    const ExcelObj& at(size_t row, size_t col) const
    {
      return *(row_begin(row) + col);
    }
    ExcelObj& at(size_t row, size_t col)
    {
      return *(row_begin(row) + col);
    }
    const ExcelObj& at(size_t n) const
    {
      const auto N = nCols();
      auto i = n / N;
      auto j = n % N;
      return at(i, j);
    }
    ExcelObj& at(size_t n)
    {
      const auto N = nCols();
      auto i = n / N;
      auto j = n % N;
      return at(i, j);
    }

    ExcelArray subArray(int fromRow, int fromCol, int toRow, int toCol)
    {
      return ExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    size_t nRows() const { return _rows; }
    size_t nCols() const { return _columns; }
    size_t size() const { return _rows * _columns; }
    size_t dims() const { return _rows > 1 && _columns > 1 ? 2 : 1; }

    const ExcelObj* row_begin(size_t i) const  { return _data + i * _baseCols + _colOffset; }
    ExcelObj* row_begin(size_t i)              { return _data + i * _baseCols + _colOffset; }
    const ExcelObj* row_end(size_t i) const    { return row_begin(i) + nCols(); }
    ExcelObj* row_end(size_t i)                { return row_begin(i) + nCols(); }

    /// <summary>
    /// Determines the type of data stored in the array if it is homogenous. If it is
    /// not, it returns the type BigData.
    ///
    /// It assumes that boolean can be interprets as integers and that integers can 
    /// be interpreted as float.  It also assumes "empty" can be interpreted as a floating
    /// point (e.g. NaN), but other error types cannot.
    ///
    /// Note that objects in Excel arrays can be one of: int, bool, double, error, string, empty.
    /// </summary>
    ExcelType dataType() const
    {
      using namespace msxll;
      int type = 0;
      for (auto i = 0; i < _rows; ++i)
        for (auto j = row_begin(i); j < row_end(i); ++j)
          type |= j->xltype;

      switch (type)
      {
      case xltypeBool:
        return ExcelType::Bool;

      case xltypeInt:
      case xltypeInt | xltypeBool:
        return ExcelType::Int;

      case xltypeNum:
      case xltypeInt | xltypeNum:
      case xltypeInt | xltypeNum | xltypeBool:
      case xltypeInt | xltypeNum | xltypeNil:
      case xltypeInt | xltypeNum | xltypeBool | xltypeNil:
        return ExcelType::Num;

      case xltypeStr:
        return ExcelType::Str;

      case xltypeErr:
        return ExcelType::Err;

      default:
        return ExcelType::BigData;
      }
    }

  private:
    size_t _rows;
    size_t _columns;
    size_t _colOffset;
    ExcelObj* _data;
    size_t _baseCols;

    void checkRange(size_t row, size_t col) const
    {
      if (row >= nRows() || col >= nCols())
        XLO_THROW("Array access ({0}, {1}) out of range ({2}, {3})", row, col, nRows(), nCols());
    }
  };

  
}