
#include <xloil/ExcelObj.h>
#include <xloil/ExcelCall.h>
#include <xloil/NumericTypeConverters.h>
#include <xloil/Throw.h>
#include <xloil/Date.h>
#include <xloil/Log.h>
#include <xloil/StringUtils.h>
#include <xlOil/ExcelRef.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/ExcelArray.h>
#include <array>
#include <algorithm>
#include <cstring>
#include <vector>
#include <string>


using std::string;
using std::wstring;
using std::vector;
using namespace msxll;
using namespace std::string_literals;

namespace xloil
{
namespace
{
  static_assert(sizeof(xloper12) == sizeof(xloil::ExcelObj));

  wchar_t* makePStringBuffer(size_t nChars)
  {
    nChars = std::min<size_t>(nChars, XL_STRING_MAX_LEN);
    auto buf = new wchar_t[nChars + 1];
    buf[0] = (wchar_t)nChars;
    return buf;
  }

  wchar_t* pascalWStringFromC(const char* cstr, size_t len)
  {
    assert(cstr);
    // This will result in a wchar buffer which may be too long
    auto pstr = makePStringBuffer(len);
    auto nChars = MultiByteToWideChar(CP_UTF8, 0, cstr, len, pstr + 1, pstr[0]);
    // nChars <= pstr[0] so cast to wchar is OK
    *pstr = (wchar_t)nChars;
    return pstr;
  }

  size_t totalStringLength(const xloper12* arr, size_t nRows, size_t nCols)
  {
    size_t total = 0;
    auto endData = arr + (nRows * nCols);
    for (; arr != endData; ++arr)
      if (arr->xltype == xltypeStr)
        total += arr->val.str[0];
    return total;
  }
}

  // TODO: https://stackoverflow.com/questions/52737760/how-to-define-string-literal-with-character-type-that-depends-on-template-parame
  const wchar_t* enumAsWCString(CellError e)
  {
    switch (e)
    {
    case CellError::Null: return L"#NULL";
    case CellError::Div0: return L"#DIV/0";
    case CellError::Value: return L"#VALUE!";
    case CellError::Ref: return L"#REF!";
    case CellError::Name: return L"#NAME?";
    case CellError::Num: return L"#NUM!";
    case CellError::NA: return L"#N/A";
    case CellError::GettingData: 
    default:
      return L"#ERR!";
    }
  }
  const wchar_t* enumAsWCString(ExcelType e)
  {
    switch (e)
    {
      case ExcelType::Num:     return L"Num";
      case ExcelType::Str :    return L"Str";
      case ExcelType::Bool:    return L"Bool";
      case ExcelType::Ref :    return L"Ref";
      case ExcelType::Err :    return L"Err";
      case ExcelType::Flow:    return L"Flow";
      case ExcelType::Multi:   return L"Multi";
      case ExcelType::Missing: return L"Missing";
      case ExcelType::Nil :    return L"Nil";
      case ExcelType::SRef:    return L"SRef";
      case ExcelType::Int :    return L"Int";
      case ExcelType::BigData: return L"BigData";
      default:
        return L"Unknown";
    }
  }


  ExcelObj::ExcelObj(ExcelType t)
  {
    switch (t)
    {
    case ExcelType::Num: val.num = 0; break;
    case ExcelType::Int: val.w = 0; break;
    case ExcelType::Bool: val.xbool = 0; break;
    case ExcelType::Str: val.str = Const::EmptyStr().val.str; break;
    case ExcelType::Err: val.err = (int)CellError::NA; break;
    case ExcelType::Multi: val.array.rows = 0; val.array.columns = 0; break;
    case ExcelType::SRef:
    case ExcelType::Flow:
    case ExcelType::BigData:
      XLO_THROW("Flow and SRef and BigData types not supported");
    }
    xltype = int(t);
  }

  void ExcelObj::createFromChars(const char* chars, size_t len)
  {
    val.str = len == 0
      ? Const::EmptyStr().val.str
      : pascalWStringFromC(chars, len);
    xltype = xltypeStr;
  }

  ExcelObj::ExcelObj(const std::tm& dt)
    : ExcelObj(excelSerialDateFromYMDHMS(
        dt.tm_year + 1900, dt.tm_mon + 1, dt.tm_mday,
        dt.tm_hour, dt.tm_min, dt.tm_sec, 0))
  {}

  ExcelObj::ExcelObj(msxll::IDSHEET sheet, const msxll::xlref12 & ref)
  {
    if (ref.rwFirst >= XL_MAX_ROWS || ref.rwFirst > ref.rwLast)
      XLO_THROW("ExcelRef out of range fromRow={0}, toRow={1}", ref.rwFirst, ref.rwLast);
    if (ref.colFirst >= XL_MAX_COLS || ref.colFirst > ref.colLast)
      XLO_THROW("ExcelRef out of range fromCol={0}, toCol={1}", ref.colFirst, ref.colLast);

    val.mref.idSheet = sheet;
    val.mref.lpmref = new msxll::XLMREF12[1];
    val.mref.lpmref->count = 1;
    val.mref.lpmref->reftbl[0] = ref;
    xltype = msxll::xltypeRef | msxll::xlbitDLLFree;
  }


  void ExcelObj::reset() noexcept
  {
    if ((xltype & xlbitXLFree) != 0)
    {
      callExcelRaw(xlFree, this, this); // arg is not really const here!
    }
    else
    {
      switch (xtype())
      {
      case xltypeStr:
        if (val.str != nullptr && val.str != Const::EmptyStr().val.str)
          delete[] val.str;
        break;

      case xltypeMulti:
        // Arrays are allocated as an array of char which contains all their strings
        // So we don't need to loop and free them individually
        delete[] (char*)(val.array.lparray);
        break;

      case xltypeBigData:
        // Only delete if count > 0. Excel uses bigdata for async return handles and sets
        // lpbData to an int handle, but leaves cbData = 0.  We never create a bigdata object
        // (other than copying async handles) so this delete should not be triggered.
        if (val.bigdata.cbData > 0)
          delete[](char*)val.bigdata.h.lpbData;
        break;

      case xltypeRef:
        delete[] (char*)val.mref.lpmref;
        break;
      }
    }
    xltype = xltypeNil;
  }

  namespace
  {
    struct Compare
    {
      template<class T> int operator()(T l, T r) const
      {
        return l < r ? -1 : (l == r ? 0 : 1);
      }
    };
    struct CompareEqual
    {
      template<class T> int operator()(T l, T r) const
      {
        return l == r ? 0 : -1;
      }
    };
  }

  template <class TCmp>
  int doCompare(
    const ExcelObj& left, 
    const ExcelObj& right, 
    bool caseSensitive,
    bool recursive) noexcept
  {
    if (&left == &right)
      return 0;

    const auto lType = left.xtype();
    const auto rType = right.xtype();
    if (lType == rType)
    {
      switch (lType)
      {
      case xltypeNum:
        return TCmp()(left.val.num, right.val.num);
      case xltypeBool:
        return TCmp()(left.val.xbool, right.val.xbool);
      case xltypeInt:
        return TCmp()(left.val.w, right.val.w);
      case xltypeErr:
        return TCmp()(left.val.err, right.val.err);
      case xltypeMissing:
      case xltypeNil:
        return 0;

      case xltypeStr:
      {
        auto lLen = left.val.str[0];
        auto rLen = right.val.str[0];
        auto len = std::min(lLen, rLen);
        auto c = caseSensitive
          ? _wcsncoll(left.val.str + 1, right.val.str + 1, len)
          : _wcsnicoll(left.val.str + 1, right.val.str + 1, len);
        return c != 0 ? c : TCmp()(lLen, rLen);
      }
      case xltypeMulti:
      {
        auto ret = TCmp()(left.val.array.columns * left.val.array.rows,
          right.val.array.columns * right.val.array.rows);
        if (ret != 0)
          return ret;
        if (!recursive)
          return 0;

        auto arrL = left.val.array.lparray;
        auto arrR = right.val.array.lparray;
        const auto end = arrL + (left.val.array.columns * left.val.array.rows);
        while (ret == 0 && arrL < end)
          ret = TCmp()((ExcelObj&)*arrL++, (ExcelObj&)*arrR++);
        return ret;
      }
      case xltypeRef:
      case xltypeSRef:
        // Case doesn't matter as we control the string representation for ranges
        return wcscmp(left.toStringRepresentation().c_str(), right.toStringRepresentation().c_str());

      default: // BigData or Flow types - not sure why you would be comparing these?!
        return 0;
      }
    }
    else
    {
      // If both types are num/int/bool we can compare as doubles
      constexpr int typeNumeric = xltypeNum | xltypeBool | xltypeInt;

      if (((lType | rType) & ~typeNumeric) == 0)
        return TCmp()(left.toDouble(), right.toDouble());

      // Errors come last
      if (((lType | rType) & xltypeErr) != 0)
        return rType == xltypeErr ? -1 : 1;

      // We want all numerics to come before string, so mask them to zero
      return (lType & ~typeNumeric) < (rType & ~typeNumeric) ? -1 : 1;
    }
  }

  bool ExcelObj::operator==(const ExcelObj& that) const
  {
    return doCompare<CompareEqual>(*this, that, true, true) == 0;
  }

  int ExcelObj::compare(
    const ExcelObj& left,
    const ExcelObj& right,
    bool caseSensitive,
    bool recursive) noexcept
  {
    return doCompare<Compare>(left, right, caseSensitive, recursive);
  }

  std::wstring ExcelObj::toString(const wchar_t* separator) const
  {
    switch (xtype())
    {
    case xltypeNum:
      return formatStr(L"%G", val.num);

    case xltypeBool:
      return wstring(val.xbool ? L"TRUE" : L"FALSE");

    case xltypeInt:
      return std::to_wstring(val.w);

    case xltypeStr:
    {
      const size_t len = val.str ? val.str[0] : 0;
      return len == 0 ? wstring() : wstring(val.str + 1, len);
    }

    case xltypeMissing:
    case xltypeNil:
      return L"";

    case xltypeErr:
      return enumAsWCString(CellError(val.err));

    case xltypeSRef:
    case xltypeRef:
      return ExcelRef(*this).value().toString(separator);

    case xltypeMulti:
    {
      ExcelArray arr(*this);
      wstring str;
      str.reserve(arr.size() * 8); // 8 is an arb choice
      if (separator)
      {
        wstring sep(separator);
        for (ExcelArray::size_type i = 0; i < arr.size(); ++i)
          str += arr(i).toString() + sep;
        if (!str.empty())
          str.erase(str.size() - sep.length());
      }
      else
        for (ExcelArray::size_type i = 0; i < arr.size(); ++i)
          str += arr(i).toString();
      return str;
    }

    default:
      return L"#???";
    }
  }
  std::wstring ExcelObj::toStringRepresentation() const noexcept
  {
    try
    {
      switch (xtype())
      {
      case xltypeSRef:
      case xltypeRef:
      {
        ExcelRef range(*this);
        return range.address();
      }
      case xltypeMulti:
        return fmt::format(L"[{0} x {1}]", val.array.rows, val.array.columns);
      default:
        return toString();
      }
    }
    catch (...)
    {
      return L"<ERROR>"s;
    }
  }
  uint16_t ExcelObj::maxStringLength() const noexcept
  {
    switch (xtype())
    {
    case xltypeInt:
    case xltypeNum:
      return 20;

    case xltypeBool:
      return 5;

    case xltypeStr:
      return val.str[0];

    case xltypeMissing:
    case xltypeNil:
      return 0;

    case xltypeErr:
      return 8;

    case xltypeSRef:
      return XL_CELL_ADDRESS_RC_MAX_LEN + XL_SHEET_NAME_MAX_LEN;

    case xltypeRef:
      return 256 + XL_CELL_ADDRESS_RC_MAX_LEN + XL_SHEET_NAME_MAX_LEN;

    case xltypeMulti:
    {
      size_t n = 0;
      auto p = val.array.lparray;
      const auto pEnd = p + (val.array.rows * val.array.columns);
      while (p < pEnd) n += ((const ExcelObj*)p++)->maxStringLength();
      return (int16_t)std::min<size_t>(USHRT_MAX, n);
    }
    default:
      return 4;
    }
  }

  bool ExcelObj::toYMD(
    int &nYear, int &nMonth, int &nDay) const noexcept
  {
    if ((xltype & (xltypeNum | xltypeInt)) == 0)
      return false;
    const auto d = toInt();
    return excelSerialDateToYMD(d, nYear, nMonth, nDay);
  }

  bool ExcelObj::toYMDHMS(
    int & nYear, int & nMonth, int & nDay,
    int & nHours, int & nMins, int & nSecs, int & uSecs) const noexcept
  {
    if ((xltype & (xltypeNum | xltypeInt)) == 0)
      return false;
    const auto d = toDouble();
    return excelSerialDatetoYMDHMS(d, nYear, nMonth, nDay, nHours, nMins, nSecs, uSecs);
  }

  bool ExcelObj::toDateTime(std::tm& result, 
    const bool coerce, const wchar_t* format) const
  {
    switch (xtype())
    {
    case xltypeNum:
    {
      int uSecs;
      result.tm_isdst = false;
      return excelSerialDatetoYMDHMS(val.num, 
        result.tm_year, result.tm_mon, result.tm_yday,
        result.tm_hour, result.tm_min, result.tm_sec, uSecs);
    }
    case xltypeInt:
    {
      result.tm_isdst = false;
      result.tm_hour = 0;
      result.tm_min = 0;
      return excelSerialDateToYMD(val.w,
        result.tm_year, result.tm_mon, result.tm_yday);
    }
    case xltypeStr:
    {
      if (!coerce)
        return false;
      return stringToDateTime(
        asPString().view(),
        result, format);
    }
    default:
      return false;
    }
  }
  bool ExcelObj::trimmedArraySize(row_t& nRows, col_t& nCols) const
  {
    if ((xtype() & xltypeMulti) == 0)
    {
      nRows = 0; nCols = 0;
      return false;
    }

    const auto start = (ExcelObj*)val.array.lparray;
    nRows = val.array.rows;
    nCols = val.array.columns;

    auto p = start + nCols * nRows - 1;

    for (; nRows > 0; --nRows)
      for (int c = (int)nCols - 1; c >= 0; --c, --p)
        if (p->isNonEmpty())
          goto StartColSearch;

  StartColSearch:
    for (; nCols > 0; --nCols)
      for (p = start + nCols - 1; p < (start + nCols * nRows); p += val.array.columns)
        if (p->isNonEmpty())
          goto SearchDone;

  SearchDone:
    return true;
  }

  void ExcelObj::overwriteComplex(ExcelObj& to, const ExcelObj& from)
  {
    switch (from.xltype & ~(xlbitXLFree | xlbitDLLFree))
    {
    case xltypeNum:
    case xltypeBool:
    case xltypeErr:
    case xltypeMissing:
    case xltypeNil:
    case xltypeInt:
    case xltypeSRef:
      (msxll::XLOPER12&)to = (const msxll::XLOPER12&)from;
      break;

    case xltypeStr:
    {
      const auto len = from.val.str[0];
#if _DEBUG
      to.val.str = new wchar_t[len + 2];
      to.val.str[len + 1] = L'\0';  // Allows debugger to read string
#else
      to.val.str = new wchar_t[len + 1];
#endif
      wmemcpy_s(to.val.str, len + 1, from.val.str, len + 1);
      to.xltype = xltypeStr;
      break;
    }
    case xltypeMulti:
    {
      const auto nRows = from.val.array.rows;
      const auto nCols = from.val.array.columns;

      const auto* pSrc = (const ExcelObj*)from.val.array.lparray;

      size_t strLength = totalStringLength(pSrc, nRows, nCols);
      ExcelArrayBuilder arr(nRows, (ExcelArray::col_t)nCols, strLength, false);

      for (auto i = 0; i < nRows; ++i)
        for (auto j = 0; j < nCols; ++j)
        {
          switch (pSrc->xltype)
          {
          case xltypeStr:
          {
            const auto len = pSrc->val.str[0];
            arr(i, j) = pSrc->asPString();
            break;
          }
          default:
            arr(i, j) = *pSrc;
          }
          ++pSrc;
        }

      // Overwrite "to"
      new (&to) ExcelObj(arr.toExcelObj());
      break;
    }

    case xltypeBigData:
    {
      auto cbData = from.val.bigdata.cbData;

      // Either it's a block of data to copy or a handle from Excel
      if (cbData > 0 && from.val.bigdata.h.lpbData)
      {
        auto pbyte = new char[cbData];
        memcpy_s(pbyte, cbData, from.val.bigdata.h.lpbData, cbData);
        to.val.bigdata.h.lpbData = (BYTE*)pbyte;
      }
      else
        to.val.bigdata.h.hdata = from.val.bigdata.h.hdata;

      to.val.bigdata.cbData = cbData;
      to.xltype = xltypeBigData;

      break;
    }

    case xltypeRef:
    {
      auto* fromMRef = from.val.mref.lpmref;
      // mref can be null after a call to xlSheetId
      auto count = fromMRef ? fromMRef->count : 0;
      if (count > 0)
      {
        auto size = sizeof(XLMREF12) + sizeof(XLREF12)*(count - 1);
        auto* newMRef = new char[size];
        memcpy_s(newMRef, size, (char*)fromMRef, size);
        to.val.mref.lpmref = (LPXLMREF12)newMRef;
      }
      else
        to.val.mref.lpmref = nullptr;
      to.val.mref.idSheet = from.val.mref.idSheet;
      to.xltype = xltypeRef;

      break;
    }
    default:
      XLO_THROW("Unhandled xltype during copy");
    }
  }

  namespace Const
  {
    const ExcelObj& Missing()
    {
      static ExcelObj obj = ExcelObj(ExcelType::Missing);
      return obj;
    }

    const ExcelObj& Error(CellError e)
    {
      static std::array<ExcelObj, _countof(theCellErrors)> cellErrors =
      {
        ExcelObj(CellError::Null),
        ExcelObj(CellError::Div0),
        ExcelObj(CellError::Value),
        ExcelObj(CellError::Ref),
        ExcelObj(CellError::Name),
        ExcelObj(CellError::Num),
        ExcelObj(CellError::NA),
        ExcelObj(CellError::GettingData)
      };
      switch (e)
      {
      case CellError::Null:        return cellErrors[0];
      case CellError::Div0:        return cellErrors[1];
      case CellError::Value:       return cellErrors[2];
      case CellError::Ref:         return cellErrors[3];
      case CellError::Name:        return cellErrors[4];
      case CellError::Num:         return cellErrors[5];
      case CellError::NA:          return cellErrors[6];
      case CellError::GettingData: return cellErrors[7];
      }
      XLO_THROW("Unexpected CellError type");
    }
    const ExcelObj& EmptyStr()
    {
      static ExcelObj obj(PString<>::steal(L"\0"));
      return obj;
    }
  }
}

namespace
{
  // Boost hash_combine
  inline void hash_combine(size_t& /*seed*/) { }

  template <typename T, typename... Rest>
  inline void hash_combine(size_t& seed, const T& v, Rest... rest) {
    std::hash<T> hasher;
    seed ^= hasher(v) + 0x9e3779b9 + (seed << 6) + (seed >> 2);
    hash_combine(seed, rest...);
  }
}

namespace std
{
  size_t hash<xloil::ExcelObj>::operator ()(const xloil::ExcelObj& value) const
  {
    switch (value.xtype())
    {
    case xltypeInt: return hash<int>()(value.val.w);
    case xltypeNum: return hash<double>()(value.val.w);
    case xltypeBool: return hash<bool>()(value.val.xbool);
    case xltypeStr: return hash<wstring_view>()(value.asPString().view());
    case xltypeMissing:
    case xltypeNil: return 0;
    case xltypeErr: return hash<int>()(value.val.err);
    case xltypeSRef:
    {
      size_t seed = 377;
      hash_combine(seed,
        value.val.sref.ref.colFirst,
        value.val.sref.ref.colLast,
        value.val.sref.ref.rwFirst,
        value.val.sref.ref.rwLast);
      return seed;
    }
    case xltypeRef:
    {
      size_t seed = 377;
      hash_combine(seed, value.val.mref.idSheet, value.val.mref.lpmref);
      return seed;
    }
    case xltypeMulti:
      return hash<void*>()(value.val.array.lparray);
    default:
      return 0;
    }
  }
}
