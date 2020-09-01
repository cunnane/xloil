#pragma once

#include "ExcelObj.h"
#include <cassert>

namespace xloil
{
  namespace detail
  {
    class ArrayBuilderAlloc
    {
    public:
      ArrayBuilderAlloc()
      {}

      // TODO: we could support resize on this class, with a small amount
      // of string fiddling 
      ArrayBuilderAlloc(size_t nObjects, size_t stringLen)
        : _nObjects(nObjects)
      {
        _buffer = (ExcelObj*)new char[sizeof(ExcelObj) * nObjects + sizeof(wchar_t) * stringLen];
        _stringData = (wchar_t*)(_buffer + nObjects);
        _endStringData = _stringData + stringLen;
      }

      PStringView<> newString(size_t len)
      {
        assert(_stringData <= _endStringData);
        _stringData[0] = wchar_t(len);
        wchar_t* ptr = _stringData;
        _stringData += len + 1;
        return PStringView<>(ptr);
      }

      ExcelObj& object(size_t i) { return _buffer[i]; }

      void fillNA()
      {
        new (_buffer) ExcelObj(CellError::NA);
        auto* source = _buffer;
        for (auto i = 1u; i < _nObjects; ++i)
          memcpy_s(_buffer + i, sizeof(ExcelObj), source, sizeof(ExcelObj));
      }

    private:
      ExcelObj* _buffer;
      size_t _nObjects;
      wchar_t* _stringData;
      const wchar_t* _endStringData;
    };

    class ArrayBuilderElement
    {
    public:
      ArrayBuilderElement(ExcelObj& target, ArrayBuilderAlloc& allocator)
        : _target(target)
        , _stringAlloc(allocator)
      {}

      template <class T, 
        std::enable_if_t<std::is_integral<T>::value, int> = 0>
      void operator=(T x) 
      { 
        // Note that _target is uninitialised memory, so we cannot 
        // call *_target = ExcelObj(x)
        new (&_target) ExcelObj(x); 
      }

      void operator=(double x) { new (&_target) ExcelObj(x); }
      void operator=(CellError x) { new (&_target) ExcelObj(x); }

      /// <summary>
      /// Move assignment from an owned pascal string.
      /// </summary>
      void operator=(PString<wchar_t>&& x) 
      { 
        new (&_target) ExcelObj(std::forward<PString<wchar_t>>(x));
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
        ExcelObj::overwrite(_target, x);
      }

      void emplace(const wchar_t* str, size_t len)
      {
        auto xlObj = new (&_target) ExcelObj();
        xlObj->xltype = msxll::xltypeStr;

        if (len == 0)
        {
          xlObj->val.str = Const::EmptyStr().val.str;
        }
        else
        {
          auto pstr = _stringAlloc.newString(len);
          wmemcpy_s(pstr.pstr(), len, str, len);
          // This object's dtor will never be called, so the allocated
          // pstr will only be freed as part of the entire array block
          xlObj->val.str = pstr.data();
        }
      }

    private:
      ExcelObj& _target;
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

      
      _allocator = detail::ArrayBuilderAlloc(arrSize, totalStrLength);
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
      return _allocator.newString(len);
    }

    /// <summary>
    /// Open a writer on the element (i, j), write to it with
    /// <code>builder(i,j) = value;</code>
    /// </summary>
    detail::ArrayBuilderElement operator()(size_t i, size_t j)
    {
      return detail::ArrayBuilderElement(element(i, j), _allocator);
    }

    detail::ArrayBuilderElement operator()(size_t i)
    {
      return detail::ArrayBuilderElement(element(i, 0), _allocator);
    }

    ExcelObj& element(size_t i, size_t j)
    {
      assert(_nRows == 1 || _nColumns == 1 || (i < _nRows && j < _nColumns));
      return _allocator.object(i * _nColumns + j);
    }

    /// <summary>
    /// Create an ExcelObj of type array from this builder. Note you
    /// can still write data using the builder after this call.
    /// </summary>
    ExcelObj toExcelObj()
    {
      return ExcelObj(&_allocator.object(0), int(_nRows), int(_nColumns));
    }

    row_t nRows() const { return _nRows; }
    col_t nCols() const { return _nColumns; }

    /// <summary>
    /// Fills the array with N/A - useful if you do not want to worry
    /// about filling in every value
    /// </summary>
    void fillNA()
    {
      _allocator.fillNA(); 
    }

  private:
    detail::ArrayBuilderAlloc _allocator;
    row_t _nRows;
    col_t _nColumns;
  };
}