#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Throw.h>
#include <cassert>

namespace xloil
{
  class ExcelArray;

  class ExcelArrayIterator
  {
  public:
    using iterator = ExcelArrayIterator;
    using col_t = ExcelObj::col_t;

    ExcelArrayIterator(
      const ExcelObj* position,
      const col_t nCols,
      const col_t baseNumCols);
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
    col_t _nCols;
    col_t _baseNumCols;
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
  /// Creates a view of an array contained in an ExcelObj. It does not 
  /// copy the array data.
  /// </summary>
  class ExcelArray
  {
  public:
    using size_type = uint32_t;
    using row_t = ExcelObj::row_t;
    using col_t = ExcelObj::col_t;

    /// <summary>
    /// Create an ExcelArray from an ExcelObj. By default trims the provided array
    /// to the last non-empty (not Nil, #N/A or "") row and column. 
    /// 
    /// Single values are converted to a 1x1 array. Throws if object type cannot 
    /// be converted to an array.
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="trim">If true, trim the array to the last non-empty row and columns</param>
    XLOIL_EXPORT explicit ExcelArray(const ExcelObj& obj, bool trim = true);

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
    ExcelArray(const ExcelArray& arr, row_t fromRow, col_t fromCol, int toRow=-1, int toCol=-1)
      : _rows((row_t)((toRow < 0 ? arr._rows + toRow + 1 : toRow) - fromRow))
      , _columns((col_t)((toCol < 0 ? arr._columns + toCol + 1 : toCol) - fromCol))
      , _baseCols(arr._baseCols)
    {
      _data = arr._data + fromRow * _baseCols + fromCol;
    }

    /// <summary>
    /// Retieves the i,j-th element from the array
    /// </summary>
    /// <param name="row"></param>
    /// <param name="col"></param>
    /// <returns>A const reference to the i,j-th element</returns>
    const ExcelObj& operator()(row_t row, col_t col) const
    {
      checkRange(row, col);
      return at(row, col);
    }

    /// <summary>
    /// Retrives the n-th element in the array, working row-wise.
    /// </summary>
    /// <param name="n"></param>
    /// <returns>A const reference to the n-th element</returns>
    const ExcelObj& operator()(size_t n) const
    {
      checkRange(n);
      return at(n);
    }

    /// <summary>
    /// Retrieves the i,j-th element without bounds checking
    /// </summary>
    /// <param name="row"></param>
    /// <param name="col"></param>
    /// <returns>A const reference to the i,j-th element</returns>
    const ExcelObj& at(row_t row, col_t col) const
    {
      return *(row_begin(row) + col);
    }
    /// <summary>
    /// Retrives the n-th element in the array, working row-wise without
    /// bounds checking.
    /// </summary>
    /// <param name="n"></param>
    /// <returns>A const reference to the n-th element</returns>
    const ExcelObj& at(size_t n) const
    {
      auto N = nCols();
      auto i = (row_t)n / N;
      auto j = (col_t)(n % N);
      return at(i, j);
    }

    /// <summary>
    /// Returns a new ExcelArray which is a sub-array of the current one.
    /// It includes fromRow and fromCol but excludes toRow and toCol.
    /// 
    /// Negative values for <paramref name="toRow"/> and <paramref name="toCol"/>
    /// are interpreted as one plus offsets from nRows and Cols respectively.
    /// </summary>
    /// <param name="fromRow"></param>
    /// <param name="fromCol"></param>
    /// <param name="toRow">The end row (not inclusive)</param>
    /// <param name="toCol">The end column (not inclusive) </param>
    /// <returns></returns>
    ExcelArray subArray(row_t fromRow, col_t fromCol, int toRow=-1, int toCol=-1) const
    {
      return ExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    row_t nRows() const { return _rows; }
    col_t nCols() const { return _columns; }
    size_type size() const { return _rows * _columns; }

    /// <summary>
    /// Returns 2 if both nRows and nCols exceed 1, otherwise returns 1
    /// </summary>
    /// <returns></returns>
    uint8_t dims() const { return _rows > 1 && _columns > 1 ? 2 : 1; }

    /// <summary>
    /// Returns an iterator to the start of the specified row. Note this iterator
    /// is only valid for the specified row, it does not wrap to the next row,
    /// use <see cref="ExcelArray::begin"/> for that functionality.
    /// </summary>
    /// <param name="i"></param>
    const ExcelObj* row_begin(row_t i) const  { return _data + i * _baseCols; }

    /// <summary>
    /// Returns an iterator to the end of the specified row (i.e. one past the last element)
    /// </summary>
    /// <param name="i"></param>
    const ExcelObj* row_end(row_t i)   const  { return row_begin(i) + nCols(); }

    /// <summary>
    /// Returns an iterator to the first element in the array
    /// </summary>
    /// <returns></returns>
    ExcelArrayIterator begin() const
    {
      return ExcelArrayIterator(row_begin(0), _columns, _baseCols);
    }

    /// <summary>
    /// Returns an iterator to the end of the array (i.e. one past the last element)
    /// </summary>
    /// <returns></returns>
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
    row_t _rows;
    col_t _columns;
    col_t _baseCols;

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
    const ExcelObj::col_t nCols,
    const ExcelObj::col_t baseNumCols)
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