#pragma once
#include "ExcelObj.h"
#include <xlOil/Throw.h>
#include <cassert>

namespace xloil
{
  class ExcelArray;

  class ExcelArrayIterator
  {
  public:
    using iterator = ExcelArrayIterator;

    ExcelArrayIterator(
      const ExcelObj* position,
      const uint16_t nCols,
      const uint16_t baseNumCols);
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
    uint16_t _nCols;
    uint16_t _baseNumCols;
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
    using size_type = uint32_t;
    using row_t = uint32_t;
    using col_t = uint16_t;

    /// <summary>
    /// Create an ExcelArray from an ExcelObj. By default trims the provided array
    /// to the last non-empty (not Nil, #N/A or "") row and column.
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="trim">If true, trim the array to the last non-empty row and columns</param>
    XLOIL_EXPORT ExcelArray(const ExcelObj& obj, bool trim = true);

    ExcelArray::ExcelArray(const ExcelObj& obj, size_t nRows, size_t nCols)
      : _rows((row_t)nRows)
      , _columns((col_t)nCols)
    {
      if (obj.type() != ExcelType::Multi)
        XLO_THROW("Expected array");
      if (nRows > (size_t)obj.val.array.rows || nCols > (size_t)obj.val.array.columns)
        XLO_THROW("Out of range");
      _data = (const ExcelObj*)obj.val.array.lparray;
      _baseCols = (col_t)obj.val.array.columns;
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
    ExcelArray(const ExcelArray& arr, size_t fromRow, size_t fromCol, int toRow=-1, int toCol=-1)
      : _rows((row_t)((toRow < 0 ? arr._rows + toRow + 1 : toRow) - fromRow))
      , _columns((col_t)((toCol < 0 ? arr._columns + toCol + 1 : toCol) - fromCol))
      , _baseCols(arr._baseCols)
    {
      _data = arr._data + fromRow * _baseCols + fromCol;
    }

    const ExcelObj& operator()(size_t row, size_t col) const
    {
      checkRange(row, col);
      return at(row, col);
    }

    const ExcelObj& operator()(size_t n) const
    {
      checkRange(n);
      return at(n);
    }

    const ExcelObj& at(size_t row, size_t col) const
    {
      return *(row_begin(row) + col);
    }

    const ExcelObj& at(size_t n) const
    {
      auto N = nCols();
      auto i = (row_t)n / N;
      auto j = (col_t)(n % N);
      return at(i, j);
    }

    ExcelArray subArray(size_t fromRow, size_t fromCol, int toRow=-1, int toCol=-1) const
    {
      return ExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    row_t nRows() const { return _rows; }
    col_t nCols() const { return _columns; }
    size_type size() const { return _rows * _columns; }
    uint8_t dims() const { return _rows > 1 && _columns > 1 ? 2 : 1; }

    const ExcelObj* row_begin(size_t i) const  { return _data + i * _baseCols; }
    const ExcelObj* row_end(size_t i)   const  { return row_begin(i) + nCols(); }

    ExcelArrayIterator begin() const
    {
      return ExcelArrayIterator(row_begin(0), _columns, _baseCols);
    }
    ExcelArrayIterator end() const 
    { 
      return ExcelArrayIterator(row_end(nRows() - 1), _columns, _baseCols);
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
      for (decltype(_rows) i = 0; i < _rows; ++i)
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
    const ExcelObj* _data;
    uint32_t _rows;
    uint16_t _columns;
    uint16_t _baseCols;

    friend class ExcelArrayIterator;

    void checkRange(size_t row, size_t col) const
    {
      if (row >= nRows() || col >= nCols())
        XLO_THROW("Array access ({0}, {1}) out of range ({2}, {3})", row, col, nRows(), nCols());
    }

    void checkRange(size_t n) const
    {
      if (n >= size())
        XLO_THROW("Array access {0} out of range {1}", n, size());
    }
  };

  inline ExcelArrayIterator::ExcelArrayIterator(
    const ExcelObj* position,
    const uint16_t nCols,
    const uint16_t baseNumCols)
    : _p(position)
    , _pRowEnd(nCols == baseNumCols ? nullptr : position + nCols)
    , _nCols(nCols)
    , _baseNumCols(baseNumCols)
  {
    // Note the optimisation: if nCols == baseNumCols, the array
    // data is contiguous so we don't need to reset the pointer
    // at the end of a row
  }

  inline ExcelArrayIterator& ExcelArrayIterator::operator++()
  {
    if (++_p == _pRowEnd)
    {
      _p += _baseNumCols - _nCols;
      _pRowEnd += _baseNumCols;
    }
    return *this;
  }
  inline ExcelArrayIterator& ExcelArrayIterator::operator--()
  {
    if (_pRowEnd - _p == _nCols)
    {
      _pRowEnd -= _baseNumCols;
      _p = _pRowEnd - 1;
    }
    else
      --_p;
    return *this;
  }
}