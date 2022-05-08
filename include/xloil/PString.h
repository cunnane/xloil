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

  template <class TChar = wchar_t, class TAlloc = PStringAllocator<TChar>> class BasicPString;
  template <class TChar = wchar_t> class BasicPStringRef;

  /// <summary>
  /// PString (length-counted string) of wide-char, owns data, behaves a 
  /// bit like `std::wstring`
  /// </summary>
  using PString     = BasicPString<wchar_t>;
  /// <summary>
  /// A non-owning reference to the data underlying a <see cref="PString"\>
  /// </summary>
  using PStringRef  = BasicPStringRef<wchar_t>;
  /// <summary>
  /// A non-owning const reference to the data underlying a <see cref="PString"\>
  /// </summary>
  using PStringCRef = BasicPStringRef<const wchar_t>;

  namespace detail
  {
    template<class T> struct StringTraits {};


    template <class TChar = wchar_t>
    class PStringImpl
    {
    public:
      using size_type = std::remove_const_t<TChar>;
      using char_type = std::remove_const_t<TChar>;
      using value_type = typename TChar;
      using const_value_type = const std::remove_const_t<value_type>;

      /// <summary>
      /// An invalid position index used to indicate "not-found", exactly
      /// as per `std::string::npos`.
      /// </summary>
      static constexpr size_type npos = size_type(-1);
      static constexpr size_t max_length = npos - 1;
      using traits = std::char_traits<char_type>;

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
      const_value_type* pstr() const { return _data + 1; }
      value_type* pstr() { return _data + 1; }

      /// <summary>
      /// Returns an iterator (really a pointer) to the beginning of the 
      /// string data.
      /// </summary>
      const_value_type* begin() const { return _data + 1; }
      value_type* begin() { return _data + 1; }

      /// <summary>
      /// Returns an iterator (really a pointer) to the end of the 
      /// string data (just past the last character).
      /// </summary>
      const_value_type* end() const { return _data + 1 + length(); }
      value_type* end() { return _data + 1 + length(); }

      /// <summary>
      /// Copy the contents of another Pascal string into this one. Throws
      /// if the existing buffer is too short.
      /// </summary>
      PStringImpl& operator=(const PStringImpl& that)
      {
        if (this != &that)
          writeOrThrow(that);
        return *this;
      }

      /// <summary>
      /// Writes the given null-terminated string into the buffer, raising an error
      /// if the buffer is too short.
      /// </summary>
      template <class T>
      PStringImpl& operator=(T str)
      {
        writeOrThrow(std::basic_string_view<char_type>(
          detail::StringTraits<T>::data(str),
          detail::StringTraits<T>::length(str)));
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

      char_type& operator[](const size_type i)
      {
        return _data[i + 1];
      }
      char_type operator[](const size_type i) const
      {
        return _data[i + 1];
      }

      operator std::basic_string_view<char_type>() const
      {
        return view();
      }

      /// <summary>
      /// Overwrites <paramref name="len"/> chars from given string into the buffer,  
      /// starting at <paramref name="start"/>. Returns true if successful or false
      /// if the internal buffer is too short.
      /// </summary>
      bool replace(size_type start, std::basic_string_view<char_type> str)
      {
        if (start + str.length() > length())
          return false;
        if (str.length() > 0)
          traits::copy(_data + 1 + start, str.data(), str.length());
        return true;
      }

      /// <summary>
      /// Overwrites <paramref name="len"/> chars from given string into the buffer,  
      /// starting at the beginning. Returns true if successful or false if the
      /// internal buffer is too short.
      /// </summary>
      bool replace(std::basic_string_view<char_type> str)
      {
        return replace(0, str);
      }

      /// <summary>
      /// Returns an STL string representation of the pascal string. This
      /// copies the string data.
      /// </summary>
      std::basic_string<char_type> string() const
      {
        return std::basic_string<char_type>(pstr(), pstr() + length());
      }

      /// <summary>
      /// Searches forward for the specified char returning its the offset
      /// of its first occurence or npos if not found.
      /// </summary>
      size_type find(char_type needle, size_type pos = 0) const
      {
        auto p = traits::find(pstr() + pos, length(), needle);
        return p ? (size_type)(p - pstr()) : npos;
      }

      /// <summary>
      /// Searches backward for the specified char returning its the offset
      /// of its last occurence or npos if not found.
      /// </summary>
      size_type rfind(char_type needle, size_type pos = npos) const
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
      std::basic_string_view<char_type> view(size_type from = 0, size_type count = npos) const
      {
        return std::basic_string_view<char_type>(
          pstr() + from, count != npos ? count : length() - from);
      }

      static char_type bound(size_t len)
      {
        return (char_type)(max_length < len ? max_length : len);
      }

      /// <summary>
      /// Like strtok but for PString. Just like strtok the source string is modified.
      /// Returns an empty BasicPStringRef when there are no more tokens.
      /// </summary>
      /// <param name="delims"></param>
      /// <returns></returns>
      BasicPStringRef<char_type> strtok(const char_type* delims);

    protected:
      TChar* _data;

      PStringImpl(TChar* data)
        : _data(data)
      {}

      void writeOrThrow(std::basic_string_view<char_type> str)
      {
        if (!replace(str))
          throw std::out_of_range(
            formatStr("PString buffer too short: %u required, %u available",
              str.length(), length()));
        _data[0] = (char_type)str.length();
      }

      void overwrite(const char_type* source, const size_type len)
      {
        traits::copy(_data + 1, source, len);
      }
    };

  } // namespace detail

  /// <summary>
  /// A Pascal string is a length-counted, rather than null-terminated string
  /// The first character in the string buffer contains its length, and the 
  /// remaining characters contain the content. This type of string is used in
  /// by Excel in its xloper (ExcelObj type in xlOil). PString helps to handle
  /// these Pascal strings by behaving somewhat like `std::string`, however it 
  /// is not recommended to use this type for generic string manipuation.
  ///  
  /// PString owns its data buffer, <see cref="BasicPStringRef"/> does not.
  /// </summary>
  template <class TChar, class TAlloc>
  class BasicPString : public detail::PStringImpl<TChar>
  {
  private:
    TAlloc _alloc;

   
  public:
    using base = detail::PStringImpl<TChar>;
    using base::size_type;
    using allocator_type = TAlloc;
    using base::_data;

    friend BasicPStringRef<TChar>;

    /// <summary>
    /// Create a PString of the specified length
    /// </summary>
    explicit BasicPString(size_type len = 0, TAlloc allocator = TAlloc())
      : base(len == 0
        ? nullptr
        : allocator.allocate((unsigned)len + 1))
      , _alloc(allocator)
    {
      if (len > 0)
        _data[0] = len;
    }

    /// <summary>
    /// Construct from an STL string_view
    /// </summary>
    /// <param name="str"></param>
    /// <param name="allocator">Optional allocator instance</param>
    explicit BasicPString(const std::basic_string_view<TChar>& str, TAlloc allocator = TAlloc())
      : BasicPString((TChar)str.length(), allocator)
    {
      overwrite(str.data(), (TChar)str.length());
    }

    /// <summary>
    /// Construct from C-string 
    /// </summary>
    /// <param name="str">C-string, must be null terminated</param>
    /// <param name="allocator">Optional allocator instance</param>
    BasicPString(const TChar* str, TAlloc allocator = TAlloc())
      : BasicPString((TChar)traits::length(str), allocator)
    {
      overwrite(str, length());
    }

    /// <summary>
    /// Construct from string literal
    /// </summary>
    /// <param name="str"></param>
    /// <param name="allocator">Optional allocator instance</param>
    template<size_t N>
    BasicPString(const TChar(*str)[N], TAlloc allocator = TAlloc())
      : BasicPString((TChar)N, allocator)
    {
      overwrite(str, (TChar)N);
    }

    /// <summary>
    /// Construct from another PString or PStringRef
    /// </summary>
    /// <param name="that"></param>
    /// <param name="allocator">Optional allocator instance</param>
    BasicPString(const base& that, TAlloc allocator = TAlloc())
      : PString(that.length(), allocator)
    {
      traits::copy(_data + 1, that.pstr(), length());
    }

    /// <summary>
    /// Move constructor
    /// </summary>
    /// <param name="that"></param>
    BasicPString(BasicPString&& that)
      : PStringImpl(nullptr)
      , _alloc(that._alloc)
    {
      std::swap(_data, that._data);
    }

    ~BasicPString()
    {
      _alloc.deallocate(_data, length() + (size_type)1);
    }

    operator PStringImpl<TChar>() const
    {
      return BasicPStringRef<TChar>(this->_data);
    }

    operator PStringImpl<TChar>()
    {
      return BasicPStringRef<TChar>(this->_data);
    }

    /// <summary>
    /// Take ownership of a Pascal string buffer, constructed externally, ideally
    /// with the same allocator
    /// </summary>
    static auto steal(TChar* data, TAlloc allocator = TAlloc())
    {
      return BasicPString(data, allocator);
    }

    BasicPString& operator=(BasicPString&& that)
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
    BasicPString& operator=(const TChar* str)
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
    BasicPString& operator=(const std::basic_string_view<TChar>& str)
    {
      const auto len = bound(str.length());
      resize(len);
      overwrite(str.data(), len);
      return *this;
    }

    using base::operator=;

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
    explicit BasicPString(TChar* data, TAlloc allocator = TAlloc())
      : base(data)
      , _alloc(allocator)
    {}
  };

  /// <summary>
  /// A view (i.e. non-owning) of a Pascal string, see the discussion in 
  /// <see cref="PString"/> for background on Pascal strings
  /// 
  /// This class cannot view a sub-string of a Pascal string, but 
  /// calling the <see cref="BasicPStringRef::view"/> method returns
  /// a std::string_view class which can.
  /// </summary>
  template <class TChar>
  class BasicPStringRef: public detail::PStringImpl<TChar>
  {
  public:
    using base = detail::PStringImpl<TChar>;
    using size_type = typename base::size_type;
    using char_type = typename base::char_type;

    /// <summary>
    /// Constructs a view of an existing Pascal string given its
    /// full data buffer (including the length count).
    /// </summary>
    explicit BasicPStringRef(TChar* data = nullptr)
      : base(data)
    {}

    /// <summary>
    /// Construct from a PString
    /// </summary>
    /// <param name="str"></param>
    BasicPStringRef(BasicPString<char_type>& str)
      : base(str._data)
    {}

    BasicPStringRef& operator=(const BasicPStringRef& that)
    {
      if (!_data)
        _data = that._data;
      else
        *(PStringImpl*)(this) = that;

      return *this;
    }

    /// <summary>
    /// Returns a pointer to the raw pascal string buffer.  Use with caution
    /// to ensure buffer lifetime is managed correctly
    /// </summary>
    auto data() const { return _data; }

    /// <summary>
    /// Resize the string buffer to the specified length. Increasing
    /// the length throws an error
    /// </summary>
    void resize(size_type sz)
    {
      if (sz <= length())
        _data[0] = sz;
      else
        throw std::out_of_range("Cannot increase size of PStringRef");
    }
  };

  template <class TChar>
  BasicPStringRef<typename detail::PStringImpl<TChar>::char_type> 
    detail::PStringImpl<TChar>::strtok(
      const typename detail::PStringImpl<TChar>::char_type* delims)
  {
    const auto nDelims = traits::length(delims);

    // If a previous PString is passed in, we will have tokenised the string
    // into [n]token[m]remaining, so the end() iterator should point to [m].
    // Otherwise we start with our own _data buffer.
    auto* p = _data;
    if (!p)
      return BasicPStringRef();

    // First character is length
    const auto stringLen = *p++;
    const auto pEnd = p + stringLen;

    // p points to the first char in the string, step until we are not
    // pointing at a delimiter. If we hit the end of the string, there
    // are no more tokens, so return a null PString.
    while (traits::find(delims, nDelims, *p))
      if (++p == pEnd)
        return BasicPStringRef();

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
    return BasicPStringRef(token - 1);
  }
  namespace detail
  {
    template <class TChar>
    BasicPString<TChar> copyRight(const TChar* that, size_t len, const PStringImpl<TChar>& right)
    {
      auto* data = BasicPString<TChar>(right.bound(right.length() + len)).release();
      std::char_traits<TChar>::copy(data + 1, that, len);
      std::char_traits<TChar>::copy(data + 1 + len, right.pstr(), right.length());
      return BasicPString<TChar>::steal(data);
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
    template<class T> struct StringTraits<BasicPStringRef<T>>
    {
      static const T* data(const PStringImpl<T>& str) {
        return str.pstr();
      }
      static size_t length(const PStringImpl<T>& str) {
        return str.length();
      }
    };
    template<class T> struct StringTraits<BasicPString<T>>
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
  inline BasicPString<TChar> operator+(const detail::PStringImpl<TChar>& left, TRight right)
  {
    const auto* r = detail::StringTraits<TRight>::data(right);
    const auto len = detail::StringTraits<TRight>::length(right);
    auto* data = BasicPString<TChar>(left.bound(left.length() + len)).release();
    std::char_traits<TChar>::copy(data + 1, left.pstr(), left.length());
    std::char_traits<TChar>::copy(data + 1 + left.length(), r, len);
    return BasicPString<TChar>::steal(data);
  }

  template <class TChar>
  inline BasicPString<TChar> operator+(const TChar* left, const detail::PStringImpl<TChar>& right)
  {
    return detail::copyRight(left, std::char_traits<TChar>::length(left), right);
  }

  template<class TChar, class _Traits, class _Alloc>
  inline std::basic_string<TChar, _Traits, _Alloc> operator+(
    const std::basic_string<TChar, _Traits, _Alloc>& left,
    const detail::PStringImpl<TChar>& right)
  {
    std::basic_string<TChar, _Traits, _Alloc> result;
    result.reserve(left.size() + right.length());
    result += left;
    result.append(right.begin(), right.end());
    return result;
  }
}