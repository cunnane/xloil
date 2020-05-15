#pragma once

#include "ExcelObj.h"
#include <xlOil/Throw.h>
#include <cassert>

namespace xloil
{
  namespace detail
  {
    class ArrayBuilderAlloc
    {
    public:
      ArrayBuilderAlloc(wchar_t* buffer, size_t bufSize)
        : _stringData(buffer)
        , _endStringData(buffer + bufSize)
      {}

      PStringView<> newString(size_t len)
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

    private:
      wchar_t* _stringData;
      const wchar_t* _endStringData;
    };

    class ArrayBuilderElement
    {
    public:
      ArrayBuilderElement(ExcelObj* target, ArrayBuilderAlloc& allocator)
        : _target(target)
        , _stringAlloc(allocator)
      {}

      template <class T, 
        std::enable_if_t<std::is_integral<T>::value, int> = 0>
      void operator=(T x) 
      { 
        // Note that _target is uninitialised memory, so we cannot 
        // call *_target = ExcelObj(x)
        new (_target) ExcelObj(x); 
      }

      void operator=(double x) { new (_target) ExcelObj(x); }
      void operator=(CellError x) { new (_target) ExcelObj(x); }

      /// <summary>
      /// Move assignment from an owned pascal string.
      /// </summary>
      void operator=(PString<wchar_t>&& x) 
      { 
        new (_target) ExcelObj(std::forward<PString<wchar_t>>(x));
      }

      /// <summary>
      /// Assign by copying data from a string_view.
      /// </summary>
      void operator=(const std::wstring_view& str)
      {
        emplace(str.data(), str.length());
      }
      void operator=(const ExcelObj& x)
      {
        assert(x.isType(ExcelType::ArrayValue));
        if (x.isType(ExcelType::Str))
        {
          auto pstr = x.asPascalStr();
          emplace(pstr.begin(), pstr.length());
        }
        else
          emplace_not_string(x);
      }

      /// <summary>
      /// Optimisation when you know the type of ExcelObj 
      /// is not a string
      /// </summary>
      void emplace_not_string(const ExcelObj& x)
      {
        assert(!x.isType(ExcelType::Str));
        ExcelObj::overwrite(*_target, x);
      }

      void emplace(const wchar_t* str, size_t len)
      {
        auto pstr = _stringAlloc.newString(len);
        wmemcpy_s(pstr.pstr(), len, str, len);
        auto xlObj = new (_target) ExcelObj();
        // We overwrite the object's string store directly, knowing its
        // d'tor will never be called as it is part of an array.
        xlObj->val.str = pstr.data();
        xlObj->xltype = msxll::xltypeStr;
      }

    private:
      ExcelObj* _target;
      ArrayBuilderAlloc& _stringAlloc;
    };
  }

  /// <summary>
  /// Constructs and allocates ExcelObj arrays. 
  /// 
  /// Usage:
  /// <code>
  ///    ExcelArrayBuilder builder(3, 1);
  ///    for (auto i = 0; i < 3; ++i)
  ///      builder(i, 0) = i;
  ///    return builder.toExcelObj();
  /// </code>
  /// </summary>
  class ExcelArrayBuilder
  {
  public:
    using row_t = ExcelObj::row_t;
    using col_t = ExcelObj::col_t;

    /// <summary>
    /// Creates an ArrayBuilder of specified size (it cannot be resized later).
    /// It does not default-initialise any ExcelObj in the array, so this must
    /// be done by the user of the class.
    /// </summary>
    /// <param name="nRows"></param>
    /// <param name="nCols"></param>
    /// <param name="totalStrLength">Total length of all strings to be added to the array</param>
    /// <param name="padTo2DimArray">Adds #N/A to ensure the array is at least 2x2</param>
    ExcelArrayBuilder(row_t nRows, col_t nCols,
      size_t totalStrLength = 0, bool padTo2DimArray = false)
      : _stringAlloc(0, 0)
    {
      // Add the terminators and string counts to total length. Maybe 
      // not every cell will be a string so this is an over-estimate
      if (totalStrLength > 0)
        totalStrLength += nCols * nRows * 2;

      auto nPaddedRows = (row_t)nRows;
      auto nPaddedCols = (col_t)nCols;
      if (padTo2DimArray)
      {
        if (nPaddedRows == 1) nPaddedRows = 2;
        if (nPaddedCols == 1) nPaddedCols = 2;
      }

      auto arrSize = nPaddedRows * nPaddedCols;

      auto* buf = new char[sizeof(ExcelObj) * arrSize + sizeof(wchar_t) * totalStrLength];
      _arrayData = (ExcelObj*)buf;
      _stringAlloc = detail::ArrayBuilderAlloc(
        (wchar_t*)(_arrayData + arrSize), totalStrLength);
      _nRows = nPaddedRows;
      _nColumns = nPaddedCols;

      if (padTo2DimArray)
      {
        // Add padding
        if (nCols < nPaddedCols)
          for (row_t i = 0; i < nRows; ++i)
            (*this)(i, nCols) = CellError::NA;

        if (nRows < nPaddedRows)
          for (col_t j = 0; j < nPaddedCols; ++j)
            (*this)(nRows, j) = CellError::NA;
      }
    }

    /// <summary>
    /// Allocate a PString in the array's string store. This is
    /// only required for some optimisations as string values
    /// assigned to ArrayBuilder elements are automatically copied
    /// into the string store.
    /// </summary>
    PStringView<> string(size_t len)
    {
      return _stringAlloc.newString(len);
    }

    /// <summary>
    /// Open a writer on the element (i, j), write to it with
    /// <code>builder(i,j) = value;</code>
    /// </summary>
    detail::ArrayBuilderElement operator()(size_t i, size_t j)
    {
      return detail::ArrayBuilderElement(element(i, j), _stringAlloc);
    }

    detail::ArrayBuilderElement operator()(size_t i)
    {
      return detail::ArrayBuilderElement(element(i, 0), _stringAlloc);
    }

    ExcelObj* element(size_t i, size_t j)
    {
      assert(_nRows == 1 || _nColumns == 1 || (i < _nRows && j < _nColumns));
      return _arrayData + (i * _nColumns + j);
    }

    /// <summary>
    /// Create an ExcelObj of type array from this builder. Note you
    /// can still write data using the builder after this call.
    /// </summary>
    ExcelObj toExcelObj()
    {
      return ExcelObj(_arrayData, int(_nRows), int(_nColumns));
    }

    row_t nRows() const { return _nRows; }
    col_t nCols() const { return _nColumns; }

  private:
    ExcelObj* _arrayData;
    detail::ArrayBuilderAlloc _stringAlloc;
    row_t _nRows;
    col_t _nColumns;
  };
}