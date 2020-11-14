#pragma once
#include <xlOil/XlCallSlim.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/PString.h>
#include <string>
#include <cassert>
#include <optional>

#define XLOIL_XLOPER msxll::xloper12

namespace xloil
{
  /// <summary>
  /// Max string length for an A1-style cell address
  /// </summary>
  constexpr uint16_t XL_CELL_ADDRESS_A1_MAX_LEN = 3 + 7 + 1 + 3 + 7 + 1;
  /// <summary>
  /// Max string length for an RC-style cell address
  /// </summary>
  constexpr uint16_t XL_CELL_ADDRESS_RC_MAX_LEN = 29;
  /// <summary>
  /// Max string length for a sheet name
  /// </summary>
  constexpr uint16_t XL_SHEET_NAME_MAX_LEN      = 31;
  /// <summary>
  /// Max string length for a (pascal) string in an ExcelObj
  /// </summary>
  constexpr uint16_t XL_STRING_MAX_LEN          = 32767;
  /// <summary>
  /// Max number of rows on a sheet
  /// </summary>
  constexpr uint32_t XL_MAX_ROWS                = 1048576;
  constexpr uint16_t XL_MAX_COLS                = 16384;
  constexpr uint16_t XL_MAX_UDF_ARGS            = 255;

  /// <summary>
  /// Describes the underlying type of the variant <see cref="ExcelObj"/>
  /// </summary>
  enum class ExcelType
  {
    Num = msxll::xltypeNum,  /// Double precision numeric data
    Str = msxll::xltypeStr,  /// Wide character string (max length 32767)
    Bool = msxll::xltypeBool, /// Boolean
    Ref = msxll::xltypeRef,  /// Range reference to one or more parts of the spreadsheet
    Err = msxll::xltypeErr,  /// Error type, <see cref="CellError"/>
    Flow = msxll::xltypeFlow, /// Legacy. Unused by xlOil.
    Multi = msxll::xltypeMulti, /// An array
    Missing = msxll::xltypeMissing, /// An omitted parameter in a function call
    Nil = msxll::xltypeNil,   /// An empty ExcelObj
    SRef = msxll::xltypeSRef,  /// Range reference to one part of the current worksheet
    Int = msxll::xltypeInt,   /// Integer type. Excel usually passes all numbers as type Num.
    BigData = msxll::xltypeStr | msxll::xltypeInt,

    /// Type group: Types that can be elements of an array. In theory nested arrays
    /// are possible by Excel will never pass one.
    ArrayValue = Num | Str | Bool | Err | Int | Nil,

    /// Type group: Types which refer to ranges
    RangeRef = SRef | Ref,

    /// Type group: Types which do not have external memory allocation
    Simple = Num | Bool | SRef | Missing | Nil | Int | Err
  };

  /// <summary>
  /// Describes the various error types Excel can handle and display
  /// </summary>
  enum class CellError
  {
    Null = msxll::xlerrNull,   /// \#NULL!
    Div0 = msxll::xlerrDiv0,   /// \#DIV0!
    Value = msxll::xlerrValue, /// \#VALUE!
    Ref = msxll::xlerrRef,     /// \#REF!
    Name = msxll::xlerrName,   /// \#NAME!
    Num = msxll::xlerrNum,     /// \#NUM!
    NA = msxll::xlerrNA,       /// \#NA!
    GettingData = msxll::xlerrGettingData /// #GETTING_DATA
  };

  /// <summary>
  /// Array of all CellError types, useful for autogeneration
  /// of bindings
  /// </summary>
  static const CellError theCellErrors[] =
  {
    CellError::Null,
    CellError::Div0,
    CellError::Value,
    CellError::Ref,
    CellError::Name,
    CellError::Num,
    CellError::NA,
    CellError::GettingData
  };

  XLOIL_EXPORT const wchar_t* enumAsWCString(CellError e);
  XLOIL_EXPORT const wchar_t* enumAsWCString(ExcelType e);

  class ExcelArray;

  class ExcelObj;
  namespace Const
  {
    /// <summary>
    /// A static ExcelObj set to Missing type.
    /// </summary>
    XLOIL_EXPORT const ExcelObj& Missing();
    /// <summary>
    /// A static ExcelObj set to the specified error type
    /// </summary>
    XLOIL_EXPORT const ExcelObj& Error(CellError e);
    /// <summary>
    /// A static ExcelObj containing an empty string
    /// </summary>
    XLOIL_EXPORT const ExcelObj& EmptyStr();
  }

  /// <summary>
  /// Wraps an XLL variant-type (xloper12) providing many useful and safe
  /// methods to create, copy and modify the data.
  /// 
  /// An ExcelObj can be statically cast to an xloper12 and vice-versa.
  /// In fact, the underlying xloper12 members can be manipulated directly,
  /// although this is not recommended.
  /// </summary>
  class XLOIL_EXPORT ExcelObj : public XLOIL_XLOPER
  {
  public:
    typedef wchar_t Char;
    typedef XLOIL_XLOPER Base;
    using row_t = uint32_t;
    using col_t = uint32_t;

    ExcelObj()
    {
      xltype = msxll::xltypeNil;
    }

    /// <summary>
    /// Constructor for integral types
    /// </summary>
    template <class T,
      std::enable_if_t<std::is_integral<T>::value, int> = 0>
    explicit ExcelObj(T x)
    {
      xltype = msxll::xltypeInt;
      val.w = (int)x;
    }

    /// <summary>
    /// Constructor for floating point types
    /// </summary>
    template <class T,
      std::enable_if_t<std::is_floating_point<T>::value, int> = 0>
    explicit ExcelObj(T d)
    {
      if (std::isnan(d))
      {
        val.err = msxll::xlerrNum;
        xltype = msxll::xltypeErr;
      }
      else
      {
        xltype = msxll::xltypeNum;
        val.num = (double)d;
      }
    }

    /// <summary>
    /// Construct from a bool
    /// </summary>
    explicit ExcelObj(bool b)
    {
      xltype = msxll::xltypeBool;
      val.xbool = b ? 1 : 0;
    }

    /// <summary>
    /// Creates an empty object of the specified type. "Empty" in this case
    /// means a sensible default depending on the data type.  For bool's it 
    /// is false, for numerics zero, for string it's the empty string, 
    /// for the error type it is \#N/A.
    /// </summary>
    /// <param name=""></param>
    ExcelObj(ExcelType);

    /// <summary>
    /// Construct from char string
    /// </summary>
    explicit ExcelObj(const char* str)
    {
      createFromChars(str, strlen(str));
    }
    explicit ExcelObj(char* str) : ExcelObj(const_cast<const char*>(str)) {}

    /// <summary>
    /// Construct from wide-char string
    /// </summary>
    explicit ExcelObj(const wchar_t* str)
      : ExcelObj(std::move(PString<>(str)))
    {}
    explicit ExcelObj(wchar_t* str) : ExcelObj(const_cast<const wchar_t*>(str)) {}

    /// <summary>
    /// Construct from STL wstring
    /// </summary>
    explicit ExcelObj(const std::wstring& s)
      : ExcelObj(std::move(PString<>(s)))
    {}

    /// <summary>
    /// Construct from STL string
    /// </summary>
    explicit ExcelObj(const std::string& s)
    {
      createFromChars(s.c_str(), s.length());
    }    
    
    /// <summary>
    /// Construct from string literal
    /// </summary>
    template<size_t N>
    explicit ExcelObj(const wchar_t(*str)[N])
      : ExcelObj(std::move(PString<>(str)))
    {}

    /// <summary>
    /// Construct from char-string literal
    /// </summary>
    template<size_t N>
    explicit ExcelObj(const char(*str)[N])
    {
      createFromChars(str, N);
    }

    /// <summary>
    /// Move ctor from owned Pascal string buffer. This takes ownership
    /// of the string buffer in the provided PString.
    /// </summary>
    /// <param name="pstr"></param>
    explicit ExcelObj(PString<Char>&& pstr)
    {
      val.str = pstr.release();
      if (!val.str)
        val.str = Const::EmptyStr().val.str;
      xltype = msxll::xltypeStr;
    }

    /// <summary>
    /// Construct from nullptr - creates an ExcelType::Missing
    /// </summary>
    /// <param name=""></param>
    ExcelObj(nullptr_t)
    {
      xltype = msxll::xltypeMissing;
    }

    /// <summary>
    /// Construct from a specified CellError 
    /// </summary>
    /// <param name="err"></param>
    ExcelObj(CellError err)
    {
      val.err = (int)err;
      xltype = msxll::xltypeErr;
    }

    /// <summary>
    /// Constructs an array from data. Takes ownership of data, which
    /// must be correctly arranged in memory. Use with caution!
    /// </summary>
    ExcelObj(const ExcelObj* data, int nRows, int nCols)
    {
      // Excel will crash if passed an empty array
      if (nRows <= 0 || nCols <= 0)
        throw std::range_error("ExcelObj: Cannot create empty array");
      val.array.rows = nRows;
      val.array.columns = nCols;
      val.array.lparray = (Base*)data;
      xltype = msxll::xltypeMulti;
    }

    /// <summary>
    /// Construct from iterable. The value_type of the iterable must be
    /// convertible to an ExcelObj using one of the other constructors
    /// </summary>
    template <class TIter>
    ExcelObj(TIter begin, TIter end);

    /// <summary>
    /// Construct from initialiser list
    /// </summary>
    template <class TIter>
    ExcelObj(std::initializer_list<TIter> vals)
      : ExcelObj(vals.begin(), vals.end())
    {}

    /// <summary>
    /// Catch constructor: purpose is to avoid pointers
    /// being auto-cast to integral / bool types
    /// </summary>
    template <class T> explicit ExcelObj(T* t) { static_assert(false); }

    /// <summary>
    /// Copy constructor
    /// </summary>
    ExcelObj(const ExcelObj& that)
    {
      overwrite(*this, that);
    }

    /// <summary>
    /// Move constructor
    /// </summary>
    ExcelObj(ExcelObj&& donor)
    {
      (Base&)*this = donor;
      // Mark donor object as empty
      donor.xltype = msxll::xltypeNil;
    }

    /// <summary>
    /// Construct local range reference (will refer to active sheet)
    /// </summary>
    /// <param name="ref"></param>
    ExcelObj(const msxll::xlref12& ref)
    {
      val.sref.ref = ref;
      xltype = msxll::xltypeSRef;
    }

    /// <summary>
    /// Construct global range reference. Sheet ID can be obtained from
    /// Excel's xlSheetId function.
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="ref"></param>
    ExcelObj(msxll::IDSHEET sheet, const msxll::xlref12& ref);

    ExcelObj::~ExcelObj()
    {
      reset();
    }

    /// <summary>
    /// Assignment from ExcelObj - performs a copy
    /// </summary>
    ExcelObj& operator=(const ExcelObj& that)
    {
      if (this == &that)
        return *this;
      copy(*this, that);
      return *this;
    }

    /// <summary>
    /// Move assignment - takes ownership of donor object's data
    /// and sets it to ExcelType::Nil.
    /// </summary>
    template <class TDonor>
    ExcelObj& operator=(TDonor&& that)
    {
      *this = std::move(ExcelObj(std::forward<TDonor>(that)));
      return *this;
    }

    ExcelObj& operator=(ExcelObj&& donor)
    {
      reset();
      (Base&)*this = donor;
      // Mark donor object as empty
      donor.xltype = msxll::xltypeNil;
      return *this;
    }

    /// <summary>
    /// Deletes object content and sets it to ExcelType::Nil
    /// </summary>
    void reset();

    /// <summary>
    /// The equality operator performs a deep comparison, recursing 
    /// into arrays and with case-sensitive string comparison. This 
    /// is different to the '<', operator which implements a cheap
    /// pointer comparison for arrays. 
    /// <seealso cref="compare"/>
    /// </summary>
    /// <param name="that"></param>
    /// <returns></returns>
    bool operator==(const ExcelObj& that) const;

    bool operator<=(const ExcelObj& that) const
    {
      return compare(*this, that, true, true) != 1;
    }

    /// <summary>
    /// Compares two ExcelObj using <see cref="compare"/> with a
    /// case insenstive string comparison and no recursion into arrays
    /// <seealso cref="compare"/>
    /// </summary>
    bool operator<(const ExcelObj& that) const
    {
      return compare(*this, that) == -1;
    }

    bool operator==(const CellError that) const
    {
      return this->xtype() == msxll::xltypeErr && val.err == (int)that;
    }

    /// <summary>
    /// Compare to string (will return false if the ExcelObj is not of 
    /// string type
    /// </summary>
    bool operator==(const std::wstring_view& that) const
    {
      return asPString() == that;
    }

    /// <summary>
    /// Compare to C-string (will return false if the ExcelObj is not of 
    /// string type
    /// </summary>
    bool operator==(const wchar_t* that) const
    {
      return asPString() == that;
    }

    /// <summary>
    /// Compare to string literal (will return false if the ExcelObj is not of 
    /// string type
    /// </summary>
    template<size_t N>
    bool operator==(const wchar_t(*that)[N]) const
    {
      return asPString() == that;
    }

    /// <summary>
    /// Compare to double (returns false is ExcelObj is not convertible to double)
    /// </summary>
    bool operator==(double that) const;

    /// <summary>
    /// Compare to int (returns false is ExcelObj is not convertible to int)
    /// </summary>
    bool operator==(int that) const;

    /// <summary>
    /// Compare to bool (returns false is ExcelObj is not convertible to bool)
    /// </summary>
    bool operator==(bool that) const;

    /// <summary>
    /// Compares two ExcelObjs. Returns -1 if left < right, 0 if left == right, else 1.
    /// 
    /// When the types of <paramref name="left"/> and <paramref name="right"/> are 
    /// the same, numeric and string types are compared in the expected way. Arrays
    /// are compared by size. Refs are compared as address strings.  BigData and 
    /// Flow types cannot be compared and return zero.
    /// 
    /// When the types differ, numeric types can still be compared but all others 
    /// are sorted by type number as repeated string conversions such as in a sort
    /// would be very expensive.
    /// </summary>
    /// <param name="left"></param>
    /// <param name="right"></param>
    /// <param name="caseSensitive">Whether string comparison is case sensitive</param>
    /// <param name="recursive">Recursively compare arrays or order them only by size</param>
    /// <returns>-1, 0 or +1</returns>
    static int compare(
      const ExcelObj& left,
      const ExcelObj& right,
      bool caseSensitive = false,
      bool recursive = false) noexcept;

    const Base* xloper() const { return this; }
    const Base* xloper() { return this; }

    /// <summary>
    /// Returns true if this ExcelObj has ExcelType::Missing type.
    /// </summary>
    bool ExcelObj::isMissing() const
    {
      return (xtype() & msxll::xltypeMissing) != 0;
    }

    /// <summary>
    /// Returns true if value is *not* one of: Missing, Nil
    /// error \#N/A or an empty string.
    /// </summary>
    /// <returns></returns>
    bool ExcelObj::isNonEmpty() const
    {
      using namespace msxll;
      switch (xtype())
      {
      case xltypeErr:
        return val.err != xlerrNA;
      case xltypeMissing:
      case xltypeNil:
        return false;
      case xltypeStr:
        return val.str[0] != L'\0';
      default:
        return true;
      }
    }
    
    /// <summary>
    /// Returns true if this ExcelObj is a \#N/A error
    /// </summary>
    /// <returns></returns>
    bool isNA() const
    {
      return isType(ExcelType::Err) && val.err == msxll::xlerrNA;
    }

    /// <summary>
    /// Get an enum describing the data contained in the ExcelObj
    /// </summary>
    /// <returns></returns>
    ExcelType ExcelObj::type() const
    {
      return ExcelType(xtype());
    }

    /// <summary>
    /// Returns true if this ExcelObj is of the specified type. This also
    /// works for compound types like ArrayValue and RangeRef that can't 
    /// be checked for by equality with type().
    /// </summary>
    bool isType(ExcelType type) const
    {
      return (xltype & (int)type) != 0;
    }

    /// <summary>
    /// Converts to a string. Attempts to stringify the various excel types.
    /// 
    /// The function recurses row-wise over Arrays and ranges and 
    /// concatenates the result. An optional separator may be given to
    /// insert between array/range entries.
    /// </summary>
    /// <param name="separator">optional separator to use for arrays</param>
    /// <returns></returns>
    std::wstring toString(const wchar_t* separator = nullptr) const;

    /// <summary>
    /// Similar to toString but more suitable for output of object 
    /// descriptions, for example in error messages. For this reason
    /// it doesn't throw but rather returns ``\<ERROR\>`` on failure.
    /// 
    /// Returns the same as toString except for arrays which yield 
    /// '[NxM]' where N and M are the number of rows and columns and 
    /// for ranges which return the range reference in the form 'Sheet!A1'.
    /// </summary>
    /// <returns></returns>
    std::wstring toStringRepresentation() const noexcept;

    /// <summary>
    /// Gives the maximum string length if toString is called on
    /// this object without actually attempting the conversion.
    /// This method is very fast except for arrays.
    /// </summary>
    uint16_t maxStringLength() const noexcept;

    /// <summary>
    /// Returns the string length if this object is a string, else zero.
    /// </summary>
    uint16_t stringLength() const
    {
      return xltype == msxll::xltypeStr ? val.str[0] : 0;
    }

    /// <summary>
    /// Returns a double from the object value type with some conversions:
    /// 
    /// int -> double
    /// bool -> 0 or 1
    /// missing/empty -> provided default
    /// 
    /// Other types will throw an exception.
    /// </summary>
    double toDouble(const std::optional<double> default = std::optional<double>()) const;

    /// <summary>
    /// Returns a int from the object value type with some conversions:
    /// 
    /// double -> int (if it can be stored as an int)
    /// bool -> 0 or 1
    /// missing/empty -> provided default
    /// 
    /// Other types will throw an exception.
    /// </summary>
    int toInt(const std::optional<int> default = std::optional<int>()) const;

    /// <summary>
    /// Returns a bool from the object value type with some conversions:
    /// 
    /// double/int -> false if exactly zero, true otherwise
    /// missing/empty -> provided default
    /// 
    /// Other types will throw an exception.
    /// </summary>
    bool toBool(const std::optional<bool> default = std::optional<bool>()) const;

    /// <summary>
    /// Returns the value in the ExcelObj assuming it is a double.
    /// If it isn't, will return nonsense (UB)
    /// </summary>
    double asDouble() const
    {
      assert(xtype() == msxll::xltypeNum);
      return val.num;
    }

    /// <summary>
    /// Returns the value in the ExcelObj assuming it is an int.
    /// If it isn't, will return nonsense (UB)
    /// </summary>
    int asInt() const
    {
      assert(xtype() == msxll::xltypeInt);
      return val.w;
    }

    /// <summary>
    /// Returns the value in the ExcelObj assuming it is a bool.
    /// If it isn't, will return nonsense (UB)
    /// </summary>
    bool asBool() const
    {
      assert(xtype() == msxll::xltypeBool);
      return val.xbool;
    }

    /// <summary>
    /// Returns the value in the ExcelObj assuming it is an array
    /// If it isn't, will return nonsense (UB)
    /// </summary>
    /// 
    const ExcelObj* asArray() const
    {
      assert(xtype() == msxll::xltypeMulti);
      return (const ExcelObj*)val.array.lparray;
    }

    /// <summary>
    /// Returns a PStringView object of the object's string data.
    /// If the object is not of string type, the resulting view 
    /// will be empty.
    /// </summary>
    PStringView<> asPString() const
    {
      return PStringView<>((xltype & msxll::xltypeStr) == 0 ? nullptr : val.str);
    }

    /// <summary>
    /// Call this on function result objects received from Excel to 
    /// declare that Excel must free them. This is automatically done
    /// by callExcel/tryCallExcel so only invoke this if you use Excel12v
    /// directly.
    /// </summary>
    /// <returns></returns>
    ExcelObj& fromExcel()
    {
      xltype |= msxll::xlbitXLFree;
      return *this;
    }

    /// <summary>
    /// Attempts to convert the ExcelObj to a Date. This will only
    /// succeed for numeric types. All function parameters are overwritten.
    /// </summary>
    /// <param name="nDay"></param>
    /// <param name="nMonth"></param>
    /// <param name="nYear"></param>
    /// <returns>true if conversion suceeds, else false</returns>
    bool toYMD(int &nYear, int &nMonth, int &nDay) const noexcept;

    /// <summary>
    /// Attempts to convert the ExcelObj to a Date/Time.  This will only
    /// succeed for numeric types. All function parameters are overwritten.
    /// </summary>
    /// <param name="nDay"></param>
    /// <param name="nMonth"></param>
    /// <param name="nYear"></param>
    /// <param name="nHours"></param>
    /// <param name="nMins"></param>
    /// <param name="nSecs"></param>
    /// <param name="uSecs"></param>
    /// <returns>true if conversion suceeds, else false</returns>
    bool toYMDHMS(int &nYear, int &nMonth, int &nDay, int& nHours,
      int& nMins, int& nSecs, int& uSecs) const noexcept;

    /// <summary>
    /// Attempts to convert the ExcelObj to a std::tm struct. If the ExcelObj
    /// is a string and coerce is true, it will attempt to parse that string
    /// using the given string format or if none was provided, all registered 
    /// string formats. 
    /// 
    /// Only relevant values in the tm struct will be populated, so you may 
    /// want to zero-initialise the struct before calling this function.
    /// </summary>
    /// <param name="datetime">The tm struct to be populated.</param>
    /// <param name="coerce">If true, attempt to parse strings</param>
    /// <param name="format">The date format for string parsing</param>
    /// <returns>true if conversion suceeds, else false</returns>
    bool toDateTime(
      std::tm& datetime, 
      const bool coerce = false, 
      const wchar_t* format = nullptr) const;

    /// <summary>
    /// Called by ExcelArray to determine the size of array data when
    /// blanks and \#N/A is ignored.
    /// </summary>
    /// <param name="nRows"></param>
    /// <param name="nCols"></param>
    /// <returns>false if object is not an array, else true</returns>
    bool trimmedArraySize(row_t& nRows, col_t& nCols) const;

    /// <summary>
    /// Destroys the target object and replaces it with the source
    /// </summary>
    /// <param name="to"></param>
    /// <param name="from"></param>
    static void copy(ExcelObj& to, const ExcelObj& from)
    {
      to.reset();
      overwrite(to, from);
    }

    /// <summary>
    /// Replaces the target object with the source without calling 
    /// the target's destructor.  Dangerous, but an optimisation for 
    /// uninitialised or non-allocating (e.g. numeric, error, nil) 
    /// target objects.
    /// </summary>
    /// <param name="to"></param>
    /// <param name="from"></param>
    static void overwrite(ExcelObj& to, const ExcelObj& from)
    {
      if (from.isType(ExcelType::Simple))
        (msxll::XLOPER12&)to = (const msxll::XLOPER12&)from;
      else
        overwriteComplex(to, from);
    }

    ExcelObj* toExcel()
    {
      xltype |= msxll::xlbitDLLFree;
      return this;
    }
  
    /// The xloper type made safe for use in switch statements by zeroing
    /// the memory control flags. Generally, prefer the type() function.
    int xtype() const
    {
      return xltype & ~(msxll::xlbitXLFree | msxll::xlbitDLLFree);
    }
   private:
    static void overwriteComplex(ExcelObj& to, const ExcelObj& from);
    void createFromChars(const char* chars, size_t len);
  };
}

namespace std {
  /// <summary>
  /// This hash does a non-recursive comparison, like operator '<'.
  /// For arrays it does not satisfy A==B => hash(A) == hash(B), but 
  /// for other types it will.
  /// </summary>
  template <>
  struct hash<xloil::ExcelObj> 
  {
    XLOIL_EXPORT size_t operator ()(const xloil::ExcelObj& value) const;
  };
}

#include <xloil/ArrayBuilder.h>
#include <xloil/NumericTypeConverters.h>

namespace xloil
{
  namespace detail
  {
    template<class T>
    inline size_t stringLength(const T&) { return 0; }
    inline size_t stringLength(const std::string_view& s) { return s.length(); }
    inline size_t stringLength(const std::wstring_view& s) { return s.length(); }
    inline size_t stringLength(const std::string& s) { return s.length(); }
    inline size_t stringLength(const std::wstring& s) { return s.length(); }
    inline size_t stringLength(const char* s) { return strlen(s); }
    inline size_t stringLength(const wchar_t* s) { return wcslen(s); }
  }
  template<class TIter> inline
  ExcelObj::ExcelObj(TIter begin, TIter end)
  {
    size_t stringLen = 0;
    ExcelObj::row_t nItems = 0;
    for (auto i = begin; i != end; ++i, ++nItems)
      stringLen += detail::stringLength(*i);

    ExcelArrayBuilder builder(nItems, 1, stringLen);
    size_t idx = 0;
    for (auto i = begin; i != end; ++i, ++idx)
      builder(idx, 0) = *i;

    xltype = msxll::xltypeNil;
    *this = builder.toExcelObj();
  }

  inline double ExcelObj::toDouble(const std::optional<double> default) const
  {
    return ToDouble()(*this, default.has_value() ? &default.value() : nullptr);
  }

  inline int ExcelObj::toInt(const std::optional<int> default) const
  {
    return ToInt()(*this, default.has_value() ? &default.value() : nullptr);
  }

  inline bool ExcelObj::toBool(const std::optional<bool> default) const
  {
    return ToBool()(*this, default.has_value() ? &default.value() : nullptr);
  }

  inline bool ExcelObj::operator==(double that) const
  {
    auto value = FromExcel<conv::ToDouble<std::optional<double>>>()(*this);
    return value == that;
  }
  inline bool ExcelObj::operator==(int that) const
  {
    auto value = FromExcel<conv::ToInt<std::optional<int>>>()(*this);
    return value == that;
  }
  inline bool ExcelObj::operator==(bool that) const
  {
    auto value = FromExcel<conv::ToBool<std::optional<bool>>>()(*this);
    return value == that;
  }
}