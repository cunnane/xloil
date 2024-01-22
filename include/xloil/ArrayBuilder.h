#pragma once

#include <xloil/ExcelObj.h>
#include <cassert>
#include <vector>

namespace xloil
{
  namespace detail
  {
    struct ArrayBuilderCharAllocator
    {
      ArrayBuilderCharAllocator(wchar_t*& data, const wchar_t* endData)
        : _stringData(data)
#ifdef _DEBUG
        , _endStringData(endData)
#endif
      {}

      wchar_t* allocate(size_t n)
      {
#ifdef _DEBUG
        if (_stringData + n > _endStringData)
          throw std::runtime_error("ExcelArrayBuilder: string data buffer exhausted");
#endif
        auto ptr = _stringData;
        _stringData += n;
        return ptr;
      }

      void deallocate(wchar_t*, size_t) { }

    private:
      wchar_t*& _stringData;
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
        , _stringData((wchar_t*)(_buffer + nObjects))
        , _endBuffer((char*)(_buffer + nObjects) + sizeof(wchar_t) * stringLen)
      {
        assert(nObjects > 0);
      }

      ~ArrayBuilderAlloc()
      {
        if (_buffer)
          delete[] (char*)_buffer;
      }

      auto charAllocator() { return ArrayBuilderCharAllocator(_stringData, (const wchar_t*)_endBuffer); }

      auto newString(size_t len)
      {
        auto ptr = charAllocator().allocate(len + 1);
        ptr[0] = wchar_t(len);
        return ptr;
      }

      bool ownsString(const wchar_t* str) const
      {
        return str >= (wchar_t*)_buffer && str <= (wchar_t*)_endBuffer;
      }

      ExcelObj& object(size_t i) { return _buffer[i]; }

      auto data() { return _buffer; }

      void fillNA()
      {
        new (_buffer) ExcelObj(CellError::NA);
        for (size_t *p = (size_t*)(_buffer + 1), *q = (size_t*)_buffer; p != (size_t*)(_buffer + _nObjects); ++p, ++q)
          *p = *q;
      }

      ExcelObj* release() 
      {
        auto buffer = _buffer;
        _buffer = nullptr;
        return buffer;
      }

    private:
      ExcelObj* _buffer;
      size_t _nObjects;
      const char* _endBuffer;
      wchar_t* _stringData;
    };

    class ArrayBuilderIterator;

    // TODO: share with SequentialArrayBuilder
    class ArrayBuilderElement
    {
    public:
      ArrayBuilderElement(size_t index, ArrayBuilderAlloc& allocator)
        : _target(&allocator.object(index))
        , _alloc(&allocator)
      {}

      ArrayBuilderElement(ExcelObj* target, ArrayBuilderAlloc& allocator)
        : _target(target)
        , _alloc(&allocator)
      {}

      template <class T,
        std::enable_if_t<std::is_integral<T>::value, int> = 0>
      auto& operator=(T x)
      {
        // Note that _target is uninitialised memory, so we cannot call 
        // *_target = ExcelObj(x)
        new (_target) ExcelObj(x);
        return *this;
      }

      auto& operator=(double x)    { new (_target) ExcelObj(x); return *this; }
      auto& operator=(CellError x) { new (_target) ExcelObj(x); return *this; }

      /// <summary>
      /// Assign by copying data from a string_view.
      /// </summary>
      auto& operator=(const std::wstring_view& str)
      {
        copy_string(str.data(), str.length());
        return *this;
      }

      /// <summary>
      /// Copy from an ExcelObj
      /// </summary>
      auto& operator=(const ExcelObj& x)
      {
        assign(x);
        return *this;
      }

      /// <summary>
      /// Copies from an ExcelObj. Optionally does not copy string data.
      /// This is safe when the parent ExcelObj will outlive this array.
      /// </summary>
      void assign(const ExcelObj& x)
      {
        if (!x.isType(ExcelType::ArrayValue))
          ExcelObj::overwrite(*_target, CellError::Value);
        else if (x.isType(ExcelType::Str) && !_alloc->ownsString(x.val.str.data))
        {
          auto pstr = x.cast<PStringRef>();
          copy_string(pstr.begin(), pstr.length());
        }
        else
          ExcelObj::overwrite(*_target, x);
      }

      operator const ExcelObj& () const { return *_target; }

      /// <summary>
      /// Move emplacement for an ExcelObj. Only safe if it is not a string or
      /// is a string allocated using the ArrayBuilder's charAllocator.
      /// </summary>
      void take(ExcelObj&& x)
      {
        if (!x.isType(ExcelType::ArrayValue))
          XLO_THROW(L"Invalid array element '{}'", x.toString());
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

      void copy_string(const wchar_t* str, size_t len)
      {
        auto xlObj = new (_target) ExcelObj();
        xlObj->xltype = msxll::xltypeStr;
        // This strings here will never be freed directly: the ExcelObj's dtor will never be 
        // called as it is an array element: the array block and its string data are freed in 
        // one call. However, we set the view flag for good practice!
        xlObj->val.str.xloil_view = true;

        if (len == 0)
        {
          xlObj->val.str.data = Const::EmptyStr().val.str.data;
        }
        else
        {
          auto pstr = _alloc->newString(len);
          wmemcpy_s(pstr + 1, len, str, len);
          xlObj->val.str.data = pstr;
        }
      }

    private:
      ExcelObj* _target;
      ArrayBuilderAlloc* _alloc;

      auto increment(int n) { _target += n; return *this; }
      friend class ArrayBuilderIterator;
    };

    class ArrayBuilderIterator
    {
    public:
      using iterator = ArrayBuilderIterator;
      using reference = ArrayBuilderElement;
      using pointer = ArrayBuilderElement*;
      using difference_type = size_t;
      using value_type = ArrayBuilderElement;
      using iterator_category = std::bidirectional_iterator_tag;

      ArrayBuilderIterator(ArrayBuilderElement&& element, int step = 1)
        : _current(element)
        , _step(step)
      {}

      auto& operator++()
      {
        _current.increment(_step);
        return *this;
      }
      auto& operator--()
      {
        _current.increment(-_step);
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

      auto operator+(const size_t n)
      {
        return iterator(ArrayBuilderElement(_current).increment(_step * (int)n), _step);
      }

      auto operator-(const size_t n)
      {
        return iterator(ArrayBuilderElement(_current).increment(-_step * (int)n), _step);
      }

      bool operator==(iterator other) const { return _current._target == other._current._target; }
      bool operator!=(iterator other) const { return !(*this == other); }

      const auto& operator*() const { return _current; }
      reference operator*() { return _current; }
      pointer operator->() { return &_current; }

    private:
      ArrayBuilderElement _current;
      int _step;
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
    XLOIL_EXPORT ExcelArrayBuilder(row_t nRows, col_t nCols,
      size_t totalStrLength = 0, bool padTo2DimArray = false);

    auto charAllocator() { return _allocator.charAllocator(); }

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

    auto row_begin(row_t i)
    {
      return detail::ArrayBuilderIterator((*this)(i, 0));
    }

    auto row_end(row_t i)
    {
      return row_begin(i) + _nColumns;
    }

    auto col_begin(col_t i)
    {
      return detail::ArrayBuilderIterator((*this)(0, i), _nColumns);
    }

    auto col_end(col_t i)
    {
      return col_begin(i) + _nRows;
    }

    /// <summary>
    /// Returns a pointer to the data of the ExcelObj array being built
    /// </summary>
    auto* data() { return &_allocator.object(0); }

  private:
    row_t _nRows;
    col_t _nColumns;
    detail::ArrayBuilderAlloc _allocator;

    static detail::ArrayBuilderAlloc initialiseAllocator(
      row_t& nRows, col_t& nCols, size_t strLength, bool padTo2DimArray);
  };

  namespace detail
  {
    /// <summary>
    /// Iterates through the strings in a contigous array of pstrings. The
    /// returned pointer is to the start of the pstring (i.e including size)
    /// </summary>
    template<class TChar>
    class PStringStackIterator
    {
    public:
      PStringStackIterator(TChar* location)
        : _current(location)
      {}
      const auto& operator*() const { return _current; }
      auto& operator++()
      {
        // Skip the length of the pstring plus 1, since we 
        // start pointed 1 before the string begins.
        _current += *_current + 1;
        return *this;
      }
    private:
      TChar* _current;
    };

    /// <summary>
    /// Allocates pstrings in a contiguous array backed by a vector.
    /// Care must be taken with this allocator as vector resizes can
    /// move the string data in memory.
    /// </summary>
    template<class TChar>
    class PStringStackAllocator : public PStringAllocator<TChar>
    {
    public:
      PStringStackAllocator(std::vector<TChar>& data)
        : _data(data)
      {}

      wchar_t* allocate(uint16_t n)
      {
        auto oldSize = _data.size();
        _data.resize(oldSize + n);
        return _data.data() + oldSize;
      }

      void deallocate(TChar*, size_t /*n*/)
      {}

    private:
      std::vector<TChar>& _data;
    };
  }

  /// <summary>
  /// An alternative array building strategy which does not require pre-calculation
  /// of the expected string length, however, values can only be added sequentially
  /// row-wise. Useful if calculation of the string length is expensive or awkward.
  /// </summary>
  class SequentialArrayBuilder
  {
  public:
    using row_t = ExcelObj::row_t;
    using col_t = ExcelObj::col_t;

    SequentialArrayBuilder(row_t nRows, col_t nCols, size_t expectedStrLength = 0)
      : _nRows(nRows)
      , _nColumns(nCols)
    {
      _objects.resize(nRows * nCols * sizeof(ExcelObj));
      _target = (ExcelObj*)_objects.data();
      _strings.reserve(expectedStrLength);
      // TODO: pad 2d? not so useful since the invention of spill
    }

    auto charAllocator() { return detail::PStringStackAllocator(_strings); }

    template <class T,
      std::enable_if_t<std::is_integral<T>::value, int> = 0>
    void emplace(T x)
    {
      emplace(ExcelObj(x));
    }

    void emplace(double x) { emplace(ExcelObj(x)); }
    void emplace(CellError x) { emplace(ExcelObj(x)); }

    /// <summary>
    /// Assign by copying data from a string_view.
    /// </summary>
    void emplace(const std::wstring_view& str)
    {
      auto N = (uint16_t)std::min<size_t>(str.length(), USHRT_MAX - 1);
      auto buffer = charAllocator().allocate(N + 1);
      buffer[0] = N;
      wmemcpy_s(buffer + 1, N, str.data(), N);
      emplaceNilString();
    }

    void emplace(ExcelObj&& obj)
    {
      if (obj.type() == ExcelType::Str &&
        !(obj.val.str.data >= _strings.data() && obj.val.str.data < _strings.data() + _strings.size()))
      {
        emplace(obj.cast<PStringRef>());
      }
      else
      {
        new (next()) ExcelObj(std::move(obj));
      }
    }

    auto stringLength() const { return _strings.size(); }
    auto nRows() const { return _nRows; }
    auto nColumns() const { return _nColumns; }

    /// <summary>
    /// Create an ExcelObj of type array from this builder.
    /// </summary>
    XLOIL_EXPORT ExcelObj toExcelObj();

    /// <summary>
    /// Copies the data in this array builder to an ArrayBuilder (original flavour)
    /// defined by iterators, should you need to do this.
    /// </summary>
    XLOIL_EXPORT void copyToBuilder(
      detail::ArrayBuilderIterator targetBegin, detail::ArrayBuilderIterator targetEnd);

  private:
    row_t _nRows;
    col_t _nColumns;
    std::vector<char> _objects;
    std::vector<wchar_t> _strings;
    ExcelObj* _target;

    void emplaceNilString()
    {
      // Write an empty ExcelObj and mark it with string type so we can find it later
      _target->xltype = msxll::xltypeStr;
      next();
    }

    void emplace(const PStringRef& pstr)
    {
      uint16_t N = pstr.length() + 1u;
      auto buffer = charAllocator().allocate(N);
      wmemcpy_s(buffer, N, pstr.data(), N);
      emplaceNilString();
    }

    ExcelObj* last()
    {
      return (ExcelObj*)(_objects.data() + _objects.size());
    }

    ExcelObj* next()
    {
      auto p = _target++;
      assert(p <= last());
      return p;
    }
  };
}