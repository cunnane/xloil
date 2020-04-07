#pragma once
#include "ExcelObj.h"
#include <cassert>
#include <xloil/Log.h>
namespace xloil
{
  class ExcelArray;

  class ExcelArrayIterator
  {
  public:
    using iterator = ExcelArrayIterator;

    ExcelArrayIterator(const ExcelArray& parent, const ExcelObj* where);
    iterator& operator++();
    iterator& operator--();
    iterator operator++(int)
    {
      iterator retval = *this;
      ++(*this);
      return retval;
    }
    iterator operator--(int)
    {
      iterator retval = *this;
      --(*this);
      return retval;
    }
    bool operator==(iterator other) const { return _p == other._p; }
    bool operator!=(iterator other) const { return !(*this == other); }
    const ExcelObj& operator*() const { return *_p; }
  private:
    const ExcelObj* _p;
    const ExcelObj* _pRowEnd;
    const ExcelArray& _obj;
  };
}

template<> struct std::iterator_traits<xloil::ExcelArrayIterator>
{
  using iterator_category = std::bidirectional_iterator_tag;
  using value_type = xloil::ExcelObj;
  using reference = const value_type&;
  using pointer = const value_type*;
  using difference_type = size_t;
};

namespace xloil
{
  /// <summary>
  /// Creates a view of an array contained in an ExcelObj.
  /// </summary>
  class ExcelArray
  {
  public:
    /// <summary>
    /// Create an ExcelArray from an ExcelObj. By default trims the provided array
    /// to the last non-empty (not Nil, #N/A or "") row and column.
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="trim">If true, trim the array to the last non-empty row and columns</param>
    ExcelArray(const ExcelObj& obj, bool trim = true);

    ExcelArray::ExcelArray(const ExcelObj& obj, size_t nRows, size_t nCols)
      : _colOffset(0)
      , _rows(nRows)
      , _columns(nCols)
    {
      if (obj.type() != ExcelType::Multi)
        XLO_THROW("Expected array");
      if (nRows > obj.val.array.rows || nCols > obj.val.array.columns)
        XLO_THROW("Out of range");
      _data = (const ExcelObj*)obj.val.array.lparray;
      _baseCols = obj.val.array.columns;
    }

    /// <summary>
    /// Creates an ExcelArray which is a subarry of a given one.
    /// Negative toRow or toCol parameters are interpreted as offsets from 
    /// the end (plus 1), hence setting fromRow=0 and toRow=-1 returns all rows. 
    /// </summary>
    /// <param name="arr">The parent array</param>
    /// <param name="fromRow">Starting row, included</param>
    /// <param name="fromCol">Starting column, included</param>
    /// <param name="toRow">Ending row, not included</param>
    /// <param name="toCol">Ending column, not included</param>
    ExcelArray(const ExcelArray& arr, int fromRow, int fromCol, int toRow=-1, int toCol=-1)
      : _rows((toRow < 0 ? arr._rows + toRow + 1 : toRow) - fromRow)
      , _columns((toCol < 0 ? arr._columns + toCol + 1 : toCol) - fromCol)
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

    const ExcelObj& operator()(size_t row) const
    {
      //TODO: checkRange(row, 0);
      return at(row);
    }

    const ExcelObj& at(size_t row, size_t col) const
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

    ExcelArray subArray(int fromRow, int fromCol, int toRow=-1, int toCol=-1) const
    {
      return ExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    size_t nRows() const { return _rows; }
    size_t nCols() const { return _columns; }
    size_t size() const { return _rows * _columns; }
    size_t dims() const { return _rows > 1 && _columns > 1 ? 2 : 1; }

    const ExcelObj* row_begin(size_t i) const  { return _data + i * _baseCols + _colOffset; }
    const ExcelObj* row_end(size_t i) const    { return row_begin(i) + nCols(); }

    ExcelArrayIterator begin() const
    {
      return ExcelArrayIterator(*this, row_begin(0));
    }
    ExcelArrayIterator end() const 
    { 
      return ExcelArrayIterator(*this, row_end(nRows() - 1));
    }

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
    const ExcelObj* _data;
    size_t _baseCols;

    friend class ExcelArrayIterator;

    void checkRange(size_t row, size_t col) const
    {
      if (row >= nRows() || col >= nCols())
        XLO_THROW("Array access ({0}, {1}) out of range ({2}, {3})", row, col, nRows(), nCols());
    }
  };

  inline ExcelArrayIterator::ExcelArrayIterator(const ExcelArray& parent, const ExcelObj* where)
    : _obj(parent)
    , _p(where)
    , _pRowEnd(where + parent.nCols())
  {}

  inline ExcelArrayIterator& ExcelArrayIterator::operator++()
  {
    if (++_p == _pRowEnd)
    {
      _p += _obj._baseCols - _obj.nCols();
      _pRowEnd += _obj._baseCols;
    }
    return *this;
  }
  inline ExcelArrayIterator& ExcelArrayIterator::operator--()
  {
    if (_pRowEnd - _p == _obj.nCols())
    {
      _pRowEnd -= _obj._baseCols;
      _p = _pRowEnd - 1;
    }
    else
      --_p;
    return *this;
  }
}