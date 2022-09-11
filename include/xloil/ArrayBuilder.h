#pragma once

#include <xloil/ExcelObj.h>
#include <cassert>

namespace xloil
{
  namespace detail
  {
    struct ArrayBuilderCharAllocator
    {
      ArrayBuilderCharAllocator()
      {}

      ArrayBuilderCharAllocator(wchar_t* data, size_t size)
        : _stringData(data)
#ifdef _DEBUG
        , _endStringData(data + size)
#endif
      {}
      constexpr wchar_t* allocate(size_t n)
      {
#ifdef _DEBUG
        if (_stringData + n > _endStringData)
          throw std::runtime_error("ExcelArrayBuilder: string data buffer exhausted");
#endif
        auto ptr = _stringData;
        _stringData += n;
        return ptr;
      }
      constexpr void deallocate(wchar_t*, size_t) { }
    private:
      wchar_t* _stringData;
#ifdef _DEBUG
      const wchar_t* _endStringData;
#endif
    };

    class ArrayBuilderAlloc
    {
    public:
      // TODO: we could support resize on this class, with a small amount
      // of string fiddling 
      ArrayBuilderAlloc(size_t nObjects, size_t stringLen)
        : _buffer((ExcelObj*)
          new char[sizeof(ExcelObj) * nObjects + sizeof(wchar_t) * stringLen])
        , _nObjects(nObjects)
        , _stringAllocator((wchar_t*)(_buffer + nObjects), stringLen)
      {
        assert(nObjects > 0);
      }

      ~ArrayBuilderAlloc()
      {
        if (_buffer)
          delete[] (char*)_buffer;
      }

      auto newString(size_t len)
      {
        auto ptr = _stringAllocator.allocate(len + 1);
        ptr[0] = wchar_t(len);
        return ptr;
      }

      ExcelObj& object(size_t i) { return _buffer[i]; }

      void fillNA()
      {
        new (_buffer) ExcelObj(CellError::NA);
        auto* source = _buffer;
        for (auto i = 1u; i < _nObjects; ++i)
          memcpy_s(_buffer + i, sizeof(ExcelObj), source, sizeof(ExcelObj));
      }

      const auto& charAllocator() const { return _stringAllocator; }

      ExcelObj* release() 
      {
        auto buffer = _buffer;
        _buffer = nullptr;
        return buffer;
      }

    private:
      ExcelObj* _buffer;
      size_t _nObjects;
      ArrayBuilderCharAllocator _stringAllocator;
    };

    class ArrayBuilderIterator;

    class ArrayBuilderElement
    {
    public:
      ArrayBuilderElement(size_t index, ArrayBuilderAlloc& allocator)
        : _target(&allocator.object(index))
        , _alloc(&allocator)
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
      /// Assign by copying data from a string_view.
      /// </summary>
      void operator=(const std::wstring_view& str)
      {
        copy_string(str.data(), str.length());
      }

      /// <summary>
      /// Copy from an ExcelObj
      /// </summary>
      void operator=(const ExcelObj& x)
      {
        assert(x.isType(ExcelType::ArrayValue));
        if (x.isType(ExcelType::Str))
        {
          auto pstr = x.cast<PStringRef>();
          copy_string(pstr.begin(), pstr.length());
        }
        else
          ExcelObj::overwrite(*_target, x);
      }

      /// <summary>
      /// Move emplacement for an ExcelObj. Only safe if it is not a string or
      /// is a string allocated using the ArrayBuilder's charAllocator.
      /// </summary>
      void emplace(ExcelObj&& x)
      {
        new (_target) ExcelObj(std::forward<ExcelObj>(x));
      }

      /// <summary>
      /// Emplacement for a static pascal string buffer - does not copy nor
      /// free the buffer.
      /// </summary>
      /// <param name="pstr"></param>
      void emplace_pstr(wchar_t* pstr)
      {
        new (_target) ExcelObj(PString::steal(pstr));
      }

      /// <summary>
      /// Optimisation of operator=. Safe when the type of ExcelObj is not
      /// a string or the parent ExcelObj will outlive the array.
      /// </summary>
      void overwrite(const ExcelObj& x)
      {
        ExcelObj::overwrite(*_target, x);
      }

      void copy_string(const wchar_t* str, size_t len)
      {
        auto xlObj = new (_target) ExcelObj();
        xlObj->xltype = msxll::xltypeStr;

        if (len == 0)
        {
          xlObj->val.str = Const::EmptyStr().val.str;
        }
        else
        {
          auto pstr = _alloc->newString(len);
          wmemcpy_s(pstr + 1, len, str, len);
          // This object's dtor will never be called, as it is an array element
          // so the allocated pstr will be freed when the entire array block is
          xlObj->val.str = pstr;
        }
      }

    private:
      ExcelObj* _target;
      ArrayBuilderAlloc* _alloc;

      friend class ArrayBuilderIterator;
    };

    class ArrayBuilderIterator
    {
    public:
      using iterator = ArrayBuilderIterator;

      ArrayBuilderIterator(ArrayBuilderElement&& element)
        : _current(element)
      {}

      auto& operator++()
      {
        ++_current._target;
        return *this;
      }
      auto& operator--()
      {
        --_current._target;
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

      bool operator==(iterator other) const { return _current._target == other._current._target; }
      bool operator!=(iterator other) const { return !(*this == other); }

      const auto& operator*() const { return _current; }
      auto& operator*() { return _current; }
      auto* operator->() { return &_current; }

    private:
      ArrayBuilderElement _current;
    };
  }

  /// <summary>
  /// Constructs and allocates ExcelObj arrays. This class does 
  /// not dynamically resize the array, you must know the size you
  /// need (and the total length of contained strings) upfront.
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

  private:
    static auto initialiseAllocator(
      row_t& nRows, col_t& nCols, size_t strLength, bool padTo2DimArray)
    {
      // Add the terminators and string counts to total length. Maybe 
      // not every cell will be a string so this is an over-estimate
      if (strLength > 0)
        strLength += nCols * nRows * 2;

      if (padTo2DimArray)
      {
        if (nRows == 1) nRows = 2;
        if (nCols == 1) nCols = 2;
      }

      auto arrSize = nRows * nCols;

      return detail::ArrayBuilderAlloc(arrSize, strLength);
    }

  public:
    /// <summary>
    /// Creates an ArrayBuilder of specified size (it cannot be resized later).
    /// It does not default-initialise any ExcelObj in the array, so this must
    /// be done by the user of the class. The fillNA() function can quickly
    /// achieve this.
    /// </summary>
    /// <param name="nRows"></param>
    /// <param name="nCols"></param>
    /// <param name="totalStrLength">Total length of all strings to be added to the array</param>
    /// <param name="padTo2DimArray">Adds # N/A to ensure the array is at least 2x2</param>
    ExcelArrayBuilder(row_t nRows, col_t nCols,
      size_t totalStrLength = 0, bool padTo2DimArray = false)
      : _nRows(nRows)
      , _nColumns(nCols)
      , _allocator(initialiseAllocator(_nRows, _nColumns, totalStrLength, padTo2DimArray))
    {
      if (padTo2DimArray)
      {
        // Add padding
        if (nCols < _nColumns)
          for (row_t i = 0; i < nRows; ++i)
            (*this)(i, nCols) = CellError::NA;

        if (nRows < _nRows)
          for (col_t j = 0; j < _nColumns; ++j)
            (*this)(nRows, j) = CellError::NA;
      }
    }

    const auto& charAllocator() const { return _allocator.charAllocator(); }

    /// <summary>
    /// Allocate a PString in the array's string store. This can be used for
    /// optimisations where a temporary string would otherwise be created in
    /// an ExcelObj.  Strings in an ExcelObj passed to ArrayBuilder elements
    /// are automatically copied into the string store.
    /// </summary>
    auto string(uint16_t len)
    {
      return BasicPString<wchar_t, detail::ArrayBuilderCharAllocator>(len, charAllocator());
    }

    /// <summary>
    /// Open a writer on the element (i, j), write to it with
    /// <code>builder(i,j) = value;</code>
    /// </summary>
    detail::ArrayBuilderElement operator()(size_t i, size_t j)
    {
      return detail::ArrayBuilderElement(i * _nColumns + j, _allocator);
    }

    detail::ArrayBuilderElement operator()(size_t i)
    {
      return detail::ArrayBuilderElement(i, _allocator);
    }

    ExcelObj& element(size_t i, size_t j)
    {
      assert(_nRows == 1 || _nColumns == 1 || (i < _nRows && j < _nColumns));
      return _allocator.object(i * _nColumns + j);
    }

    /// <summary>
    /// Create an ExcelObj of type array from this builder. This releases control
    /// of the data block and invalidates the builder.
    /// </summary>
    ExcelObj toExcelObj()
    {
      return ExcelObj(_allocator.release(), int(_nRows), int(_nColumns));
    }

    operator ExcelObj() { return toExcelObj(); }

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

    auto begin()
    {
      return detail::ArrayBuilderIterator((*this)(0));
    }

    auto end()
    {
      return detail::ArrayBuilderIterator((*this)(_nRows, _nColumns));
    }

    /// <summary>
    /// Returns a pointer to the data of the ExcelObj array being built
    /// </summary>
    auto* data() { return &_allocator.object(0); }

  private:
    row_t _nRows;
    col_t _nColumns;
    detail::ArrayBuilderAlloc _allocator;
  };
}