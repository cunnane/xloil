#pragma once
// TODO: avoid pulling windows and excel into every header!
#include "XlCallSlim.h"
#include <string>
#include <array>

#define XLOIL_XLOPER msxll::xloper12

namespace xloil
{
  enum class ExcelType
  {
    Num     = 0x0001,
    Str     = 0x0002,
    Bool    = 0x0004,
    Ref     = 0x0008,
    Err     = 0x0010,
    Flow    = 0x0020,
    Multi   = 0x0040,
    Missing = 0x0080,
    Nil     = 0x0100,
    SRef    = 0x0400,
    Int     = 0x0800,
    BigData = msxll::xltypeStr | msxll::xltypeInt
  };

  enum class CellError
  {
    Null  = msxll::xlerrNull,
    Div0  = msxll::xlerrDiv0,
    Value = msxll::xlerrValue,
    Ref   = msxll::xlerrRef,
    Name  = msxll::xlerrName,
    Num   = msxll::xlerrNum,
    NA    = msxll::xlerrNA,
    GettingData = msxll::xlerrGettingData
  };

  static const std::array<CellError, 8> theCellErrors =
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

  //const char* toCString(ExcelError e);
  const wchar_t* toWCString(CellError e);


  class ExcelArray;

  template <class TChar>
  class PString
  {
  public:
    PString(const TChar* pascalStr) : _data(pascalStr) {}

    size_t size() const { return _data[0]; }
    const TChar* buf() const { return _data; }
    std::wstring string() const { return std::wstring(_data, _data + _data[0]); }

  private:
    const TChar* _data;
  };

  class ExcelObj;
  namespace Const
  {
    const ExcelObj& Missing();
    const ExcelObj& Error(CellError e);
    const ExcelObj& EmptyStr();
  }

  class ExcelObj : public XLOIL_XLOPER
  {
  public:
    typedef wchar_t Char;
    typedef XLOIL_XLOPER Base;

    // If you got here, consider casting. It's main purpose is to avoid
    // pointers being auto-cast to integral / bool types
    template <class T> ExcelObj(T t) { static_assert(false); }

    ExcelObj();

    // Whole bunch of numeric types auto-casted to make life easier
    explicit ExcelObj(int);
    explicit ExcelObj(unsigned int x) : ExcelObj((int)x) {}
    explicit ExcelObj(long x) : ExcelObj((int)x) {}
    explicit ExcelObj(unsigned long x) : ExcelObj((int)x) {}
    explicit ExcelObj(unsigned short x) : ExcelObj((int)x) {}
    explicit ExcelObj(short x) : ExcelObj((int)x) {}
    explicit ExcelObj(double);
    explicit ExcelObj(float x) : ExcelObj((double)x) {}
    explicit ExcelObj(bool);

    /// <summary>
    /// Creates an empty object of the specified type. "Empty" has a sensible 
    /// default depending on the data type. For the  error type it is #N/A.
    /// </summary>
    /// <param name=""></param>
    ExcelObj(ExcelType);

    //template <class TAlloc> ExcelObj(const char*, TAlloc = AllocNew());
    ExcelObj(const char*);
    ExcelObj(const wchar_t*);
    ExcelObj(const std::string&);
    ExcelObj(const std::wstring&);
    ExcelObj(nullptr_t);
    ExcelObj(const ExcelObj&);
    ExcelObj(ExcelObj&&);
    ExcelObj(CellError err);

    /// Constructs an array from data without copying it. Do not use
    // TODO: declare these private and use friend? 
    ExcelObj(const ExcelObj* data, int nRows, int nCols);

    /// Non copying ctor from pascal string buffer. Do not use
    ExcelObj(const PString<Char>& pstr);


    /// <summary>
    /// Constructs a string of size nChars, returning a pointer to the internal buffer
    /// nChars may be altererd to reflect Excel's max string length.
    /// This constructor is inteded to avoid a string copy
    /// </summary>
    /// <param name="buf"></param>
    /// <param name="nChars">requested buffer size, will be capped at Excel's limit</param>
    ExcelObj(wchar_t*& buf, size_t& nChars);


    /// <summary>
    /// String constructor to guarantee copy elison via RVO, so you can write 
    /// return ExcelObj(...)
    /// </summary>
    template <class TStrWrite>
    ExcelObj(size_t nChars, TStrWrite fn, wchar_t* do_not_use_buf=nullptr)
      : ExcelObj(do_not_use_buf, nChars)
    {
      fn(do_not_use_buf, nChars);
    }

    //template<class T> ExcelObj(const T* array, int nRows, int nCols)
    //  : ExcelObj(emplace(array, nRows, nCols), nRows, nCols)
    //{
    //}

    ~ExcelObj();

    /// <summary>
    /// Deletes object content and sets it to #N/A
    /// </summary>
    void reset();

    ExcelObj& operator=(const ExcelObj& that);
    ExcelObj& operator=(ExcelObj&& that);

    bool operator==(const ExcelObj& that) const;
    static int compare(const ExcelObj& left, const ExcelObj& right);

    const Base* ptr() const { return this; }
    const Base* cptr() const { return this; }
    Base* ptr() { return this; }

    bool isMissing() const;
    /// <summary>
    /// Returns true if value is not one of: type missing, type nil
    /// error #N/A or empty string.
    /// </summary>
    /// <returns></returns>
    bool isNonEmpty() const;

    /// <summary>
    /// Get an enum describing the data contained in the ExcelObj
    /// </summary>
    /// <returns></returns>
    ExcelType type() const;


    /// <summary>
    /// Converts to a string. If strict=false then attempts
    /// to stringify the various excel types
    /// </summary>
    /// <param name="strict"></param>
    /// <returns></returns>
    std::wstring toString(bool strict = false) const;
    double toDouble() const;
    int toInt() const;
    bool toBool() const;


    double asDouble() const;
    int asInt() const;
    bool asBool() const;

    /*template <class T> T to() const {}
    template <> int to<int>() const { return toInt(); }
    template <> double to<double>() const { return toDouble(); }
    template <> std::wstring to<std::wstring>() const { return toString(); }
    template <> bool to<bool>() const { return toBool(); }*/

    /// <summary>
    /// Gets the start of the object's string data and the length
    /// of that string. The string may not be null-terminated.
    /// </summary>
    /// <param name="length"></param>
    /// <returns>Pointer to string data or null if not a string</returns>
    const wchar_t* asPascalStr(size_t& length) const;

    /// <summary>
    /// Writes the characters in supplied buffer to object's internal
    /// string buffer. Returns zero if object is not of string type.
    /// </summary>
    /// <param name="buf"></param>
    /// <param name="bufSize">Number of characters written</param>
    /// <returns></returns>
    size_t writeString(wchar_t* buf, size_t bufSize) const;

    static void copy(ExcelObj& to, const ExcelObj& from);

    /// <summary>
    /// Call this on function result objects received from Excel to 
    /// declare that Excel must free them. This is automatically done
    /// by callExcel/tryCallExcel so only invoke this if you use Excel12v
    /// directly.
    /// </summary>
    /// <returns></returns>
    ExcelObj& fromExcel();

    /// <summary>
    /// Returns a pointer to the current object suitable for returning to Excel
    /// as the result of a function. Modifies the object to tell Excel that we
    /// must free the memory via the xlAutoFree callback. Only use this on the 
    /// final object which is passed back to Excel.
    /// </summary>
    /// <returns></returns>
    ExcelObj* toExcel();

    template<class... Args>
    static ExcelObj* returnValue(Args&&... args)
    {
      return (new ExcelObj(std::forward<Args>(args)...))->toExcel();
    }
    static ExcelObj* returnValue(CellError err)
    {
      return const_cast<ExcelObj*>(&Const::Error(err));
    }
    template<>
    static ExcelObj* returnValue(ExcelObj&& p)
    {
      // same as the args, but want to make this one explicit
      return (new ExcelObj(std::forward<ExcelObj>(p)))->toExcel();
    }

    // TODO: implement coercion from string
    bool toDMY(int &nDay, int &nMonth, int &nYear, bool coerce = false);
    bool toDMYHMS(int &nDay, int &nMonth, int &nYear, int& nHours, int& nMins, int& nSecs, int& uSecs, bool coerce = false);

    /// <summary>
    /// Called by ExcelArray to determine the size of array data when
    /// blanks and #N/A is ignored.
    /// </summary>
    /// <param name="nRows"></param>
    /// <param name="nCols"></param>
    /// <returns>false if object is not an array, else true</returns>
    bool trimmedArraySize(int& nRows, int& nCols) const;

  private:
    /// The xloper type made safe for use in switch statements by zeroing
    /// the memory control flags blanked.
    int xtype() const;
  };

  size_t xlrefToString(const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize);

}
