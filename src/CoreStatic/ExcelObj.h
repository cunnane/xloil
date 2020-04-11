#pragma once
// TODO: avoid pulling windows and excel into every header!
#include "XlCallSlim.h"
#include "PString.h"
#include <string>
#include <array>
#include <cassert>

#define XLOIL_XLOPER msxll::xloper12

#define XLO_RETURN_ERROR(err) return ExcelObj::returnValue(err)

namespace xloil
{
  constexpr size_t CELL_ADDRESS_A1_MAX_LEN = 3 + 7 + 1 + 3 + 7 + 1;
  constexpr size_t CELL_ADDRESS_RC_MAX_LEN = 29;
  constexpr size_t WORKSHEET_NAME_MAX_LEN = 31;
  constexpr size_t XL_STR_MAX_LEN = 32767;

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
    BigData = msxll::xltypeStr | msxll::xltypeInt,
    ArrayValue = Num | Str | Bool | Err | Int | Nil
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

  const wchar_t* enumAsWCString(CellError e);
  const wchar_t* enumAsWCString(ExcelType e);

  class ExcelArray;

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

    // If you got here, consider casting. Its main purpose is to avoid
    // pointers being auto-cast to integral / bool types
    template <class T> ExcelObj(T t) { static_assert(false); }

    ExcelObj()
    {
      xltype = msxll::xltypeNil;
    }

    // Whole bunch of numeric types auto-casted to make life easier
    explicit ExcelObj(int);
    explicit ExcelObj(unsigned int x) : ExcelObj((int)x) {}
    explicit ExcelObj(long x) : ExcelObj((int)x) {}
    explicit ExcelObj(long long x) : ExcelObj((int)x) {}
    explicit ExcelObj(unsigned long x) : ExcelObj((int)x) {}
    explicit ExcelObj(unsigned short x) : ExcelObj((int)x) {}
    explicit ExcelObj(short x) : ExcelObj((int)x) {}
    explicit ExcelObj(double);
    explicit ExcelObj(float x) : ExcelObj((double)x) {}
    explicit ExcelObj(bool);

    /// <summary>
    /// Creates an empty object of the specified type. "Empty" in this case
    /// means a sensible default depending on the data type.  For bool's it 
    /// is false, for numerics zero, for string it's the empty string, 
    /// for the error type it is #N/A.
    /// </summary>
    /// <param name=""></param>
    ExcelObj(ExcelType);

    ExcelObj(const char*);
    ExcelObj(const wchar_t*);
    
    ExcelObj::ExcelObj(const std::string& s)
      :ExcelObj(s.c_str())
    {}

    ExcelObj::ExcelObj(const std::wstring& s)
      : ExcelObj(s.c_str())
    {}

    ExcelObj::ExcelObj(nullptr_t)
    {
      xltype = msxll::xltypeMissing;
    }

    ExcelObj(const ExcelObj&);

    ExcelObj::ExcelObj(ExcelObj&& that)
    {
      // Steal all data
      this->val = that.val;
      this->xltype = that.xltype;
      // Mark donor object as empty
      that.xltype = msxll::xltypeNil;
    }

    ExcelObj::ExcelObj(CellError err)
    {
      val.err = (int)err;
      xltype = msxll::xltypeErr;
    }

    /// Constructs an array from data without copying it. Do not use
    // TODO: declare these private and use friend? 
    ExcelObj(const ExcelObj* data, int nRows, int nCols);

    /// Non copying ctor from pascal string buffer.
    ExcelObj(PString<Char>&& pstr);
    ExcelObj(PString<Char>& pstr) : ExcelObj(std::forward<PString<Char>>(pstr)) {}
  
    ExcelObj::~ExcelObj()
    {
      reset();
    }

    /// <summary>
    /// Deletes object content and sets it to #N/A
    /// </summary>
    void reset();

    ExcelObj& operator=(const ExcelObj& that);
    ExcelObj& operator=(ExcelObj&& that);

    bool operator==(const ExcelObj& that) const
    {
      return compare(*this, that) == 0;
    }

    static int compare(const ExcelObj& left, const ExcelObj& right);

    const Base* ptr() const { return this; }
    const Base* cptr() const { return this; }
    Base* ptr() { return this; }

    bool ExcelObj::isMissing() const
    {
      return (xtype() &  msxll::xltypeMissing) != 0;
    }
    /// <summary>
    /// Returns true if value is not one of: type missing, type nil
    /// error #N/A or empty string.
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
    /// Get an enum describing the data contained in the ExcelObj
    /// </summary>
    /// <returns></returns>
    ExcelType ExcelObj::type() const
    {
      return ExcelType(xtype());
    }

    bool isRangeRef() const 
    {
      return (xltype & (msxll::xltypeRef | msxll::xltypeSRef)) != 0;
    }

    /// <summary>
    /// Converts to a string. If strict=false then attempts
    /// to stringify the various excel types, otherwise will
    /// give an error for non string types.
    /// 
    /// The function recurses row-wise over Arrays and ranges and 
    /// concatenates the result.
    /// </summary>
    /// <param name="strict"></param>
    /// <returns></returns>
    std::wstring toString(bool strict = false) const;

    /// <summary>
    /// Similar to toString but more suitable for output of object 
    /// descriptions, for example in error messages.  
    /// 
    /// Returns the same as toString except for arrays which yield 
    /// '[NxM]' where N and M are the number of rows and columns and 
    /// for ranges which return the range reference in the form 'Sheet!A1'.
    /// </summary>
    /// <returns></returns>
    std::wstring toStringRepresentation() const;

    /// <summary>
    /// Gives the maximum string length if toString is called on
    /// this object without actually attempting the conversion.
    /// </summary>
    size_t maxStringLength() const;

    /// <summary>
    /// Returns the string length if this object is a string, else zero.
    /// </summary>
    size_t stringLength() const
    {
      return xltype == msxll::xltypeStr ? val.str[0] : 0;
    }

    double ExcelObj::toDouble() const;

    int ExcelObj::toInt() const;

    bool ExcelObj::toBool() const;

    double ExcelObj::asDouble() const
    {
      assert(xtype() == msxll::xltypeNum);
      return val.num;
    }

    int ExcelObj::asInt() const
    {
      assert(xtype() == msxll::xltypeInt);
      return val.w;
    }

    bool ExcelObj::asBool() const
    {
      assert(xtype() == msxll::xltypeBool);
      return val.xbool;
    }

    const ExcelObj* ExcelObj::asArray() const
    {
      assert(xtype() == msxll::xltypeMulti);
      return (const ExcelObj*)val.array.lparray;
    }

    /// <summary>
    /// Returns a PStringView object of the object's string data.
    /// If the object is not of string type, the resulting view 
    /// will be empty.
    /// </summary>
    PStringView<> ExcelObj::asPascalStr() const
    {
      if ((xtype() & msxll::xltypeStr) == 0)
        return PStringView<>();
      return PStringView<>(val.str);
    }


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
    static ExcelObj* returnValue(const std::exception& e)
    {
      return returnValue(e.what());
    }

    template<>
    static ExcelObj* returnValue(ExcelObj&& p)
    {
      // same as the args, but want to make this one explicit
      return (new ExcelObj(std::forward<ExcelObj>(p)))->toExcel();
    }

    // TODO: implement coercion from string
    bool toDMY(int &nDay, int &nMonth, int &nYear, bool coerce = false);
    bool toDMYHMS(int &nDay, int &nMonth, int &nYear, int& nHours, 
      int& nMins, int& nSecs, int& uSecs, bool coerce = false);

    /// <summary>
    /// Called by ExcelArray to determine the size of array data when
    /// blanks and #N/A is ignored.
    /// </summary>
    /// <param name="nRows"></param>
    /// <param name="nCols"></param>
    /// <returns>false if object is not an array, else true</returns>
    bool trimmedArraySize(uint32_t& nRows, uint16_t& nCols) const;

  private:
    /// The xloper type made safe for use in switch statements by zeroing
    /// the memory control flags blanked.
    int ExcelObj::xtype() const
    {
      return xltype & ~(msxll::xlbitXLFree | msxll::xlbitDLLFree);
    }
  };

  size_t xlrefToStringRC(const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize);
}
