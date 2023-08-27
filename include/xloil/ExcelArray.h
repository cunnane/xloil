#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Throw.h>
#include <cassert>

namespace xloil
{
  namespace detail
  {
    template <class TSize>
    inline bool sliceIndices(int& from, int to, TSize& size)
    {
      if (from < 0)
      {
        from += size;
        if (from < 0)
          return false;
      }
      if (to < 0)
      {
        to += size;
        if (to < 0)
          return false;
      }
      else if (to > (int)size)
        return false;
      const auto sz = to - from;
      if (sz < 0)
        return false;
      size = (TSize)sz;
      return true;
    }
  }

  class ExcelArray;

  class ExcelArrayIterator
  {
  public:
    using iterator = ExcelArrayIterator;
    using col_t = ExcelObj::col_t;

    ExcelArrayIterator(
      const ExcelObj* position,
      const col_t nCols,
      const col_t baseNumCols) noexcept;
    iterator& operator++() noexcept;
    iterator& operator--() noexcept;
    iterator operator++(int) noexcept
    {
      iterator retval = *this;
      ++(*this);
      return retval;
    }
    iterator operator--(int) noexcept
    {
      iterator retval = *this;
      --(*this);
      return retval;
    }
    bool operator==(iterator other) const noexcept { return _p == other._p; }
    bool operator!=(iterator other) const noexcept { return !(*this == other); }
    const ExcelObj& operator*() const noexcept { return *_p; }
    auto* operator->() noexcept { return &*_p; }

  private:
    const ExcelObj* _p;
    const ExcelObj* _pRowEnd;
    col_t _nCols;
    col_t _baseNumCols;
  };


  /// <summary>
  /// Yes I know there's one in boost and there's ranges::stride_view in C++23 but we're
  /// in 2022 and I don't want to include boost to save 20 lines!
  /// </summary>
  template<class T>
  class StrideIterator
  {
  public:
    using iterator = StrideIterator<T>;

    StrideIterator(T start, ptrdiff_t stride)
      : _p(start)
      , _stride(stride)
    {}
    auto& operator++()
    {
      _p += _stride;
      return *this;
    }
    auto& operator--()
    {
      _p -= _stride;
      return *this;
    }
    auto operator++(int)
    {
      iterator copy = *this;
      ++(*this);
      return copy;
    }
    auto operator--(int)
    {
      iterator copy = *this;
      --(*this);
      return copy;
    }

    bool operator==(iterator other) const { return _p == other._p; }
    bool operator!=(iterator other) const { return !(*this == other); }

    const auto& operator*() const { return *_p; }
    auto& operator*() { return *_p; }
    auto* operator->() { return &*_p; }

  private:
    T _p;
    ptrdiff_t _stride;
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

template<class T> struct std::iterator_traits<xloil::StrideIterator<T>>
{
  using iterator_category = std::bidirectional_iterator_tag;
  using value_type = typename std::iterator_traits<T>::value_type;
  using reference = typename std::iterator_traits<T>::reference;
  using pointer = typename std::iterator_traits<T>::pointer;
  using difference_type = size_t;
};

namespace xloil
{
  /// <summary>
  /// Creates a view of an array contained in an ExcelObj. It does not 
  /// copy or own the array data.
  /// </summary>
  class ExcelArray
  {
  public:
    using size_type = uint32_t;
    using row_t = ExcelObj::row_t;
    using col_t = ExcelObj::col_t;

    /// <summary>
    /// Create an ExcelArray from an ExcelObj. By default trims the provided array
    /// to the last non-empty (not Nil, \#N/A or "") row and column. 
    /// 
    /// Single values are converted to a 1x1 array. Throws if object type cannot 
    /// be converted to an array.
    /// </summary>
    /// <param name="obj"></param>
    /// <param name="trim">If true, trim the array to the last non-empty row and columns</param>
    XLOIL_EXPORT explicit ExcelArray(const ExcelObj& obj, bool trim = true);

    /// <summary>
    /// Creates an ExcelArray which is a subarray of a given one. It extends from 
    /// (fromRow, fromCol) to (toRow, toCol) not including the right-hand ends.
    /// 
    /// Negative values for the parameters are interpreted as offsets from nRows and 
    /// nCols respectively.
    /// </summary>
    /// <param name="arr">The parent array</param>
    /// <param name="fromRow">Starting row, included</param>
    /// <param name="fromCol">Starting column, included</param>
    /// <param name="toRow">Ending row, not included</param>
    /// <param name="toCol">Ending column, not included</param>
    ExcelArray(const ExcelArray& arr,
      int fromRow, int fromCol,
      int toRow, int toCol)
      : _baseCols(arr._baseCols)
      , _rows(arr.nRows())
      , _columns(arr.nCols())
    {
      if (!detail::sliceIndices(fromRow, toRow, _rows))
        XLO_THROW_TYPE(std::out_of_range, "Invalid sub-array row indices {0}, {1} in array of size ({2}, {3})",
          fromRow, toRow, arr.nRows());
      if (!detail::sliceIndices(fromCol, toCol, _columns))
        XLO_THROW_TYPE(std::out_of_range, "Invalid sub-array column indices {0}, {1} in array of size ({2}, {3})",
          fromCol, toCol, arr.nCols());

      _data = arr._data + (col_t)fromRow * _baseCols + fromCol;
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
    /// Retrives the n-th element in the array, working row-wise. Rather slow so use sparingly.
    /// </summary>
    /// <returns>A const reference to the n-th element</returns>
    const ExcelObj& operator()(size_t n) const
    {
      checkRange(n);
      return at(n);
    }

    const ExcelObj& operator[](size_t n) const
    {
      checkRange(n);
      return at(n);
    }

    /// <summary>
    /// Retrieves the i,j-th element without bounds checking
    /// </summary>
    /// <returns>A const reference to the i,j-th element</returns>
    const ExcelObj& at(row_t row, col_t col) const
    {
      return *(row_begin(row) + col);
    }
    /// <summary>
    /// Retrives the n-th element in the array, working row-wise without bounds checking.
    /// Rather slow so use sparingly.
    /// </summary>
    /// <returns>A const reference to the n-th element</returns>
    const ExcelObj& at(size_t n) const
    {
      auto N = nCols();
      auto i = (row_t)n / N;
      auto j = (col_t)(n % N);
      return at(i, j);
    }

    /// <summary>
    /// Returns a sub-array as a new ExcelArray.
    /// It extends from (fromRow, fromCol) to the end of the array.
    /// 
    /// Negative values for <paramref name="fromRow"/> and <paramref name="fromCol"/>
    /// are interpreted as offsets from nRows and nCols respectively.
    /// </summary>
    /// <param name="fromRow"></param>
    /// <param name="fromCol"></param>
    /// <returns></returns>
    ExcelArray slice(int fromRow, int fromCol) const
    {
      return ExcelArray(*this, fromRow, fromCol, nRows(), nCols());
    }
    /// <summary>
    /// Returns a sub-array as a new ExcelArray.
    /// It extends from (fromRow, fromCol) to (toRow, toCol) not including the right
    /// hand ends.
    /// 
    /// Negative values for the parameters are interpreted as offsets from nRows and 
    /// nCols respectively.
    /// </summary>
    /// <returns></returns>
    ExcelArray slice(
      int fromRow, int fromCol,
      int toRow, int toCol) const
    {
      return ExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    row_t nRows() const { return _rows; }
    col_t nCols() const { return _columns; }
    size_type size() const { return _rows * _columns; }

    /// <summary>
    /// Returns 2 if both nRows and nCols exceed 1, otherwise returns 1
    /// unless the array has zero size, in which case returns zero.
    /// </summary>
    /// <returns></returns>
    uint8_t dims() const
    {
      return _rows > 1 && _columns > 1
        ? 2
        : (_rows == 0 || _columns == 0 ? 0 : 1);
    }

    /// <summary>
    /// Returns an iterator to the start of the specified row. Note this iterator
    /// is only valid for the specified row, it does not wrap to the next row,
    /// use <see cref="ExcelArray::begin"/> for that functionality.
    /// </summary>
    const ExcelObj* row_begin(row_t i) const { return _data + i * (size_t)_baseCols; }

    /// <summary>
    /// Returns an iterator to the end of the specified row (i.e. one past the last element)
    /// </summary>
    /// <param name="i"></param>
    const ExcelObj* row_end(row_t i)   const { return row_begin(i) + nCols(); }

    /// <summary>
    /// Returns an iterator to the start of the specified column. Does not check that
    /// `j` is a valid column.
    /// </summary>
    auto col_begin(col_t j) const
    {
      assert(j < _columns);
      return StrideIterator<const ExcelObj*>(_data + j, _baseCols);
    }

    /// <summary>
    /// Returns an iterator to the end of the specified column. Does not check that
    /// `j` is a valid column.
    /// </summary>
    auto col_end(col_t j) const
    {
      assert(j < _columns);
      return StrideIterator<const ExcelObj*>(_data + j + (_rows - 1) * _baseCols, _baseCols);
    }

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
      // The whole array iterator steps beyond the end of the last start of the
      // next row which may be different if _columns != _baseCols
      return ExcelArrayIterator(row_begin(_columns > 0 ? nRows() : 0), _columns, _baseCols);
    }

    /// <summary>
    /// Determines the type of data stored in the array if it is homogenous. If it is
    /// not, it returns the type BigData.
    ///
    /// It assumes that boolean can be interpreted as integers and that integers can 
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

    /// <summary>
    /// Returns an ExcelObj of Array type which contains the array data
    /// viewed by this ExcelArray. The data is copied. 
    /// you
    /// </summary>
    XLOIL_EXPORT ExcelObj toExcelObj() const;

    /// <summary>
    /// Determine the size of array data when blanks and \#N/A is ignored.
    /// </summary>
    /// <returns>false if object is not an array, else true</returns>
    XLOIL_EXPORT static bool trimmedArraySize(const ExcelObj& obj, row_t& nRows, col_t& nCols);

    /// <summary>
    /// Returns an ExcelObj of Array type which contains the array data.
    /// Tries to avoid copying the underlying data where possible. This
    /// function should only be used when the underlying data is guaranteed
    /// to outlive the object returned from this function.
    /// </summary>
    ExcelObj toExcelObjUnsafe() const
    {
      if (dims() == 0)
        return ExcelObj();

      if (_columns == _baseCols || _rows == 1)
        return ExcelObj(_data, _rows, _columns, true);
      else
      {
        auto data = new ExcelObj::Base[size()];
        // The iterator is noexcept so the raw ptr above is safe
        std::copy(begin(), end(), data);
        return ExcelObj((ExcelObj*)data, _rows, _columns);
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
        XLO_THROW_TYPE(std::out_of_range, "Array access ({0}, {1}) out of range ({2}, {3})", row, col, nRows(), nCols());
    }

    void checkRange(size_t n) const
    {
      if (n >= size())
        XLO_THROW_TYPE(std::out_of_range, "Array access {0} out of range {1}", n, size());
    }
  };

  inline ExcelArrayIterator::ExcelArrayIterator(
    const ExcelObj* position,
    const ExcelObj::col_t nCols,
    const ExcelObj::col_t baseNumCols) noexcept
    : _p(position)
    , _pRowEnd(nCols == baseNumCols ? nullptr : position + nCols)
    , _nCols(nCols)
    , _baseNumCols(baseNumCols)
  {
    // Note the optimisation: if nCols == baseNumCols, the array
    // data is contiguous so we don't need to reset the pointer
    // at the end of a row
  }

  inline ExcelArrayIterator& ExcelArrayIterator::operator++() noexcept
  {
    if (++_p == _pRowEnd)
    {
      _p += _baseNumCols - _nCols;
      _pRowEnd += _baseNumCols;
    }
    return *this;
  }
  inline ExcelArrayIterator& ExcelArrayIterator::operator--() noexcept
  {
    if (_pRowEnd - _p == (ptrdiff_t)_nCols)
    {
      _pRowEnd -= _baseNumCols;
      _p = _pRowEnd - 1;
    }
    else
      --_p;
    return *this;
  }
}