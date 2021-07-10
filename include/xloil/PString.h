#pragma once
#include <xloil/StringUtils.h>
#include <string>

namespace xloil
{
  /// <summary>
  /// Searches backward for the specified char returning a pointer
  /// to its last occurence or null if not found. Essentialy it is
  /// wmemchr backwards.
  /// </summary>
  inline const wchar_t*
    wmemrchr(const wchar_t* ptr, wchar_t wc, size_t num)
  {
    for (; num; --ptr, --num)
      if (*ptr == wc)
        return ptr;
    return nullptr;
  }

  template<class T>
  struct PStringAllocator
  {
    constexpr T* allocate(size_t n)
    {
      return new T[n];
    }
    constexpr void deallocate(T* p, size_t /*n*/)
    {
      delete[] p;
    }
  };

  template <class TChar = wchar_t, class TAlloc = PStringAllocator<TChar>> class PString;
  template <class TChar = wchar_t> class PStringView;
  namespace detail
  {
    template<class T> struct StringTraits {};
  }

  template <class TChar = wchar_t>
  class PStringImpl
  {
  public:
    using size_type = TChar;
    static constexpr size_type npos = size_type(-1);
    static constexpr size_t max_length = (TChar)-1 - 1;
    using traits = std::char_traits<TChar>;


    /// <summary>
    /// Returns true if the string is empty
    /// </summary>
    bool empty() const { return !_data || _data[0] == 0; }

    /// <summary>
    /// Returns the length of the string. The length is limited by
    /// sizeof(TChar).
    /// </summary>
    size_type length() const { return _data ? _data[0] : 0; }

    /// <summary>
    /// Returns a pointer to the start of the string data similar to 
    /// string::c_str, however, the string data is not guaranteed to be
    /// null-terminated.
    /// </summary>
    const TChar* pstr() const { return _data + 1; }
    TChar* pstr() { return _data + 1; }

    /// <summary>
    /// Returns an iterator (really a pointer) to the beginning of the 
    /// string data.
    /// </summary>
    const TChar* begin() const { return _data + 1; }
    TChar* begin() { return _data + 1; }

    /// <summary>
    /// Returns an iterator (really a pointer) to the end of the 
    /// string data (just past the last character).
    /// </summary>
    const TChar* end() const { return _data + 1 + length(); }
    TChar* end() { return _data + 1 + length(); }

    /// <summary>
    /// Copy the contents of another Pascal string into this one. Throws
    /// if the existing buffer is too short.
    /// </summary>
    PStringImpl& operator=(const PStringImpl& that)
    {
      if (this != &that)
        writeOrThrow(that.pstr(), that.length());
      return *this;
    }

    /// <summary>
    /// Writes the given null-terminated string into the buffer, raising an error
    /// if the buffer is too short.
    /// </summary>
    template <class T>
    PStringImpl& operator=(T str)
    {
      writeOrThrow(
        detail::StringTraits<T>::data(str),
        detail::StringTraits<T>::length(str));
      return *this;
    }

    template <class T>
    bool operator==(T that) const
    {
      return view() == that;
    }

    template <class T>
    bool operator!=(const T& that)
    {
      return !(*this == that);
    }

    wchar_t& operator[](const size_type i)
    {
      return _data[i + 1];
    }
    wchar_t operator[](const size_type i) const
    {
      return _data[i + 1];
    }

    operator std::basic_string_view<TChar>() const
    {
      return view();
    }

    /// <summary>
    /// Overwrites <paramref name="len"/> chars from given string into the buffer,  
    /// starting at <paramref name="start"/>. Returns true if successful or false
    /// if the internal buffer is too short.
    /// </summary>
    bool replace(TChar start, size_t len, const TChar* str)
    {
      if (start + len > length())
        return false;
      if (len > 0)
        traits::copy(_data + 1 + start, str, len);
      return true;
    }

    /// <summary>
    /// Overwrites <paramref name="len"/> chars from given string into the buffer,  
    /// starting at the beginning. Returns true if successful or false if the
    /// internal buffer is too short.
    /// </summary>
    bool replace(size_t len, const TChar* str)
    {
      return replace(0, len, str);
    }

    /// <summary>
    /// Returns an STL string representation of the pascal string. This
    /// copies the string data.
    /// </summary>
    std::basic_string<TChar> string() const
    {
      return std::basic_string<TChar>(pstr(), pstr() + length());
    }

    /// <summary>
    /// Searches forward for the specified char returning its the offset
    /// of its first occurence or npos if not found.
    /// </summary>
    size_type find(TChar needle, size_type pos = 0) const
    {
      auto p = traits::find(pstr() + pos, length(), needle);
      return p ? (size_type)(p - pstr()) : npos;
    }

    /// <summary>
    /// Searches backward for the specified char returning its the offset
    /// of its last occurence or npos if not found.
    /// </summary>
    size_type rfind(TChar needle, size_type pos = npos) const
    {
      auto p = wmemrchr(
        pstr() + (pos == npos ? length() : pos), 
        needle, 
        length() - (pos == npos ? 0 : pos));
      return p ? (size_type)(p - pstr()) : npos;
    }

    /// <summary>
    /// Returns a STL string_view of the string data or, optionally,
    /// a substring of it.
    /// </summary>
    std::basic_string_view<TChar> view(size_type from = 0, size_type count = npos) const
    {
      return std::basic_string_view<TChar>(
        pstr() + from, count != npos ? count : length() - from);
    }

    static TChar bound(size_t len)
    {
      return (TChar)(max_length < len ? max_length : len);
    }

  protected:
    TChar* _data;

    PStringImpl(TChar* data)
      : _data(data)
    {}

    void writeOrThrow(const TChar* str, size_t len)
    {
      if (!replace(len, str))
        throw std::out_of_range(
          formatStr("PString buffer too short: %u required, %u available",
            len, length()));
      _data[0] = (TChar)len;
    }

    void overwrite(const TChar* source, TChar len)
    {
      traits::copy(_data + 1, source, len);
    }
  };

  /// <summary>
  /// A Pascal string is a length-counted, rather than null-terminated string
  /// The first character in the string buffer contains its length, and the 
  /// remaining characters contain the content. This type of string is used in
  /// by Excel in its xloper (ExcelObj type in xlOil). PString helps to handle
  /// these Pascal strings by behaving somewhat like std::string, however it 
  /// is not recommended to use this type for generic string manipuation.
  ///  
  /// PString owns its data buffer, <see cref="PStringView"/> does not.
  /// </summary>
  template <class TChar, class TAlloc>
  class PString : public PStringImpl<TChar>
  {
  private:
    TAlloc _alloc;

  public:
    using size_type = PStringImpl::size_type;
    using allocator_type = TAlloc;

    friend PStringView<TChar>;

    /// <summary>
    /// Create a PString of the specified length
    /// </summary>
    explicit PString(size_type length = 0, TAlloc allocator = TAlloc())
      : PStringImpl(length == 0
        ? nullptr
        : allocator.allocate(length + 1))
      , _alloc(allocator)
    {
      if (length > 0)
        _data[0] = length;
    }

    /// <summary>
    /// Construct from an STL string_view
    /// </summary>
    /// <param name="str"></param>
    /// <param name="allocator">Optional allocator instance</param>
    explicit PString(const std::basic_string_view<TChar>& str, TAlloc allocator = TAlloc())
      : PString((TChar)str.length(), allocator)
    {
      overwrite(str.data(), (TChar)str.length());
    }

    /// <summary>
    /// Construct from C-string 
    /// </summary>
    /// <param name="str">C-string, must be null terminated</param>
    /// <param name="allocator">Optional allocator instance</param>
    PString(const TChar* str, TAlloc allocator = TAlloc())
      : PString((TChar)traits::length(str), allocator)
    {
      overwrite(str, length());
    }

    /// <summary>
    /// Construct from string literal
    /// </summary>
    /// <param name="str"></param>
    /// <param name="allocator">Optional allocator instance</param>
    template<size_t N>
    PString(const TChar(*str)[N], TAlloc allocator = TAlloc())
      : PString((TChar)N, allocator)
    {
      overwrite(str, (TChar)N);
    }

    /// <summary>
    /// Construct from another PString or PStringView
    /// </summary>
    /// <param name="that"></param>
    /// <param name="allocator">Optional allocator instance</param>
    PString(const PStringImpl& that, TAlloc allocator = TAlloc())
      : PString(that.length(), allocator)
    {
      traits::copy(_data + 1, that.pstr(), length());
    }

    /// <summary>
    /// Move constructor
    /// </summary>
    /// <param name="that"></param>
    PString(PString&& that)
      : PStringImpl(nullptr)
      , _alloc(that._alloc)
    {
      std::swap(_data, that._data);
    }

    ~PString()
    {
      _alloc.deallocate(_data, length() + (size_type)1);
    }

    /// <summary>
    /// Take ownership of a Pascal string buffer, constructed externally, ideally
    /// with the same allocator
    /// </summary>
    static auto steal(TChar* data, TAlloc allocator = TAlloc())
    {
      return PString(data, allocator);
    }

    PString& operator=(PString&& that)
    {
      _alloc.deallocate(_data, length() + 1);
      _data = nullptr;
      std::swap(_data, that._data);
      _alloc = that._alloc;
      return *this;
    }

    /// <summary>
  /// Writes the given null-terminated string into the buffer, raising an error
  /// if the buffer is too short.
  /// </summary>
    PStringImpl& operator=(const TChar* str)
    {
      const auto len = bound(traits::length(str));
      resize(len);
      overwrite(str, len);
      return *this;
    }
    /// <summary>
    /// Writes the given string_view into the buffer, raising an error
    /// if the buffer is too short.
    /// </summary>
    PStringImpl& operator=(const std::basic_string_view<TChar>& str)
    {
      const auto len = bound(str.length());
      resize(len);
      overwrite(str.data(), len);
      return *this;
    }

    using PStringImpl::operator=;

    /// <summary>
    /// Returns a pointer to the buffer containing the string and
    /// relinquishes ownership.
    /// </summary>
    TChar* release()
    {
      TChar* d = nullptr;
      std::swap(d, _data);
      return d;
    }

    /// <summary>
    /// Resize the string buffer to the specified length. Increasing
    /// the length forces a string copy.
    /// </summary>
    void resize(size_type sz)
    {
      if (sz <= length())
        _data[0] = sz;
      else
      {
        PString copy(sz);
        traits::copy(copy._data + 1, _data, sz);
        *this = std::move(copy);
      }
    }
  private:
    explicit PString(TChar* data, TAlloc allocator = TAlloc())
      : PStringImpl(data)
      , _alloc(allocator)
    {}
  };

  /// <summary>
  /// A view (i.e. non-owning) of a Pascal string, see the discussion in 
  /// <see cref="PString"/> for background on Pascal strings
  /// 
  /// This class cannot view a sub-string of a Pascal string, but 
  /// calling the <see cref="PStringView::view"/> method returns
  /// a std::string_view class which can.
  /// </summary>
  template <class TChar>
  class PStringView : public PStringImpl<TChar>
  {
  public:
    /// <summary>
    /// Constructs a view of an existing Pascal string given its
    /// full data buffer (including the length count).
    /// </summary>
    explicit PStringView(TChar* data = nullptr)
      : PStringImpl(data)
    {}

    /// <summary>
    /// Construct from a PString
    /// </summary>
    /// <param name="str"></param>
    PStringView(PString<TChar>& str)
      : PStringImpl(str._data)
    {}

    PStringImpl& operator=(const PStringView& that)
    {
      if (!_data)
      {
        _data = that._data;
        return *this;
      }
      else
        return *(PStringImpl*)(this) = that;
    }

    /// <summary>
    /// Returns a pointer to the raw pascal string buffer.  Use with caution
    /// to ensure buffer lifetime is managed correctly
    /// </summary>
    TChar* data() const { return _data; }

    /// <summary>
    /// Resize the string buffer to the specified length. Increasing
    /// the length throws an error
    /// </summary>
    void resize(size_type sz)
    {
      if (sz <= length())
        _data[0] = sz;
      else
        throw std::out_of_range("Cannot increase size of PStringView");
    }

    /// <summary>
    /// Like strtok but for PString. Just like strtok the source string is modified.
    /// Returns an empty PStringView when there are no more tokens.
    /// </summary>
    /// <param name="delims"></param>
    /// <returns></returns>
    PStringView strtok(const TChar* delims)
    {
      const auto nDelims = traits::length(delims);

      // If a previous PString is passed in, we will have tokenised the string
      // into [n]token[m]remaining, so the end() iterator should point to [m].
      // Otherwise we start with our own _data buffer.
      auto* p = _data;
      if (!p)
        return PStringView();

      // First character is length
      const auto stringLen = *p++;
      const auto pEnd = p + stringLen;

      // p points to the first char in the string, step until we are not
      // pointing at a delimiter. If we hit the end of the string, there
      // are no more tokens, so return a null PString.
      while (traits::find(delims, nDelims, *p))
        if (++p == pEnd)
          return PStringView();

      // p now points the first non-delimiter, the start of our token
      auto* token = p;

      // Find the next delimiter
      while (p < pEnd && !traits::find(delims, nDelims, *p)) ++p;
      const auto tokenLen = (TChar)(p - token);

      // We know token[-1] must point to a delimiter or a length count,
      // so it is safe to overwrite with the token length
      token[-1] = tokenLen;

      // If there still more string, overwrite p (which points to a delimiter)
      // with the remaining length for subsequent calls to strtok.
      if (p < pEnd)
      {
        *p = (TChar)(pEnd - p - 1);
        _data = p;
      }
      else
        _data = nullptr;
      return PStringView(token - 1);
    }
  };

  namespace detail
  {
    template <class TChar>
    PString<TChar> copyRight(const TChar* that, size_t len, const PStringImpl<TChar>& right)
    {
      auto* data = PString<TChar>(right.bound(right.length() + len)).release();
      std::char_traits<TChar>::copy(data + 1, that, len);
      std::char_traits<TChar>::copy(data + 1 + len, right.pstr(), right.length());
      return PString<TChar>::steal(data);
    }
    template<class T> struct StringTraits<std::basic_string_view<T>>
    {
      static const T* data(const std::basic_string_view<T>& str) {
        return str.data();
      }
      static size_t length(const std::basic_string_view<T>& str) {
        return str.length();
      }
    };
    template<class T> struct StringTraits<std::basic_string<T>>
    {
      static const T* data(const std::basic_string<T>& str) {
        return str.c_str();
      }
      static size_t length(const std::basic_string<T>& str) {
        return str.length();
      }
    };
    template<class T> struct StringTraits<T*>
    {
      static const T* data(const T* str) {
        return str;
      }
      static size_t length(const T* str) {
        return std::char_traits<T>::length(str);
      }
    };
    template<class T, size_t N> struct StringTraits<T(*)[N]>
    {
      static const T* data(const T(*str)[N]) {
        return str;
      }
      static size_t length(const T(*str)[N]) {
        return N;
      }
    };
    template<class T> struct StringTraits<PStringView<T>>
    {
      static const T* data(const PStringImpl<T>& str) {
        return str.pstr();
      }
      static size_t length(const PStringImpl<T>& str) {
        return str.length();
      }
    };
    template<class T> struct StringTraits<PString<T>>
    {
      static const T* data(const PStringImpl<T>& str) {
        return str.pstr();
      }
      static size_t length(const PStringImpl<T>& str) {
        return str.length();
      }
    };
  }

  template <class TChar, class TRight>
  inline PString<TChar> operator+(const PStringImpl<TChar>& left, TRight right)
  {
    const auto* r = detail::StringTraits<TRight>::data(right);
    const auto len = detail::StringTraits<TRight>::length(right);
    auto* data = PString<TChar>(left.bound(left.length() + len)).release();
    std::char_traits<TChar>::copy(data + 1, left.pstr(), left.length());
    std::char_traits<TChar>::copy(data + 1 + left.length(), r, len);
    return PString<TChar>::steal(data);
  }

  template <class TChar>
  inline PString<TChar> operator+(const TChar* left, const PStringImpl<TChar>& right)
  {
    return detail::copyRight(left, std::char_traits<TChar>::length(left), right);
  }

  template<class TChar, class _Traits, class _Alloc>
  inline std::basic_string<TChar, _Traits, _Alloc> operator+(
    const std::basic_string<TChar, _Traits, _Alloc>& left,
    const PStringImpl<TChar>& right)
  {
    std::basic_string<TChar, _Traits, _Alloc> result;
    result.reserve(left.size() + right.length());
    result += left;
    result.append(right.begin(), right.end());
    return result;
  }
}