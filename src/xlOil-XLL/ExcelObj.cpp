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
#include <xloil/StringUtils.h>
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
    // The wchar buffer may be too long if there are multibyte chars in the input
    auto pstr = makePStringBuffer(len);
    auto nChars = MultiByteToWideChar(CP_UTF8, 0, cstr, int(len), pstr + 1, pstr[0]);
    // nChars <= pstr[0] so cast to wchar is OK
    *pstr = (wchar_t)nChars;
    if (nChars == 0 && GetLastError() == ERROR_INSUFFICIENT_BUFFER)
    {
      nChars = MultiByteToWideChar(
        CP_UTF8, 0, cstr, std::min<int>(nChars, XL_STRING_MAX_LEN), pstr + 1, pstr[0]);
    }
    return pstr;
  }

  size_t totalStringLength(const xloper12* arr, size_t nRows, size_t nCols)
  {
    size_t total = 0;
    auto endData = arr + (nRows * nCols);
    for (; arr != endData; ++arr)
      if (arr->xltype == xltypeStr)
        total += arr->val.str.data[0];
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
    case CellError::GettingData: return L"#GETTING_DATA";
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
    val.str.data = len == 0
      ? Const::EmptyStr().val.str.data
      : pascalWStringFromC(chars, len);
    val.str.xloil_view = false;
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
        if (val.str.data != nullptr && val.str.data != Const::EmptyStr().val.str.data
          && !val.str.xloil_view)
          PStringAllocator<wchar_t>().deallocate(val.str.data, 0);
        break;

      case xltypeMulti:
        // Arrays are allocated as an array of char which contains all their strings
        // So we don't need to loop and free them individually. If we are at this point
        // we must have created the ExcelObj ourselves, so it is safe to use the
        // xloil_view extension.
        if (!val.array.xloil_view)
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
        auto lLen = left.val.str.data[0];
        auto rLen = right.val.str.data[0];
        auto len = std::min(lLen, rLen);
        auto c = caseSensitive
          ? _wcsncoll(left.val.str.data + 1, right.val.str.data + 1, len)
          : _wcsnicoll(left.val.str.data + 1, right.val.str.data + 1, len);
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
        return wcscmp(left.toString().c_str(), right.toString().c_str());

      default: // BigData or Flow types - not sure why you would be comparing these?!
        return 0;
      }
    }
    else
    {
      // If both types are num/int/bool we can compare as doubles
      constexpr int typeNumeric = xltypeNum | xltypeBool | xltypeInt;

      if (((lType | rType) & ~typeNumeric) == 0)
        return TCmp()(left.get<double>(), right.get<double>());

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

  std::wstring ExcelObj::toStringRecursive(const wchar_t* separator) const
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
      const size_t len = val.str.data ? val.str.data[0] : 0;
      return len == 0 ? wstring() : wstring(val.str.data + 1, len);
    }

    case xltypeMissing:
    case xltypeNil:
      return L"";

    case xltypeErr:
      return enumAsWCString(CellError(val.err));

    case xltypeSRef:
    case xltypeRef:
      return ExcelRef(*this).value().toStringRecursive(separator);

    case xltypeMulti:
    {
      ExcelArray arr(*this); // Note that this trims the array
      wstring str;
      str.reserve(arr.size() * 8); // 8 is an arb choice
      if (separator)
      {
        wstring sep(separator);
        for (ExcelArray::size_type i = 0; i < arr.size(); ++i)
          str += arr(i).toStringRecursive() + sep;
        if (!str.empty())
          str.erase(str.size() - sep.length());
      }
      else
        for (ExcelArray::size_type i = 0; i < arr.size(); ++i)
          str += arr(i).toStringRecursive();
      return str;
    }

    default:
      return L"#???";
    }
  }
  std::wstring ExcelObj::toString() const noexcept
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
        return formatStr(L"[%d x %d]", val.array.rows, val.array.columns);
      default:
        return toStringRecursive();
      }
    }
    catch (...)
    {
      // TODO: when does this ever happen?
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
      return val.str.data[0];

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
      auto p = (const ExcelObj*)val.array.lparray;
      const auto pEnd = p + (val.array.rows * val.array.columns);
      while (p < pEnd) n += (p++)->maxStringLength();
      return (int16_t)std::min<size_t>(USHRT_MAX, n);
    }
    default:
      return 4;
    }
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
      const auto len = from.val.str.data[0];
      // preserve view?
#if _DEBUG
      to.val.str.data = new wchar_t[len + 2];
      to.val.str.data[len + 1] = L'\0';  // Allows debugger to read string
#else
      to.val.str.data = new wchar_t[len + 1];
#endif
      wmemcpy_s(to.val.str.data, len + 1, from.val.str.data, len + 1);
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
            arr(i, j) = pSrc->cast<PStringRef>();
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
    namespace
    {
      static ExcelObj theMissing(ExcelType::Missing);

      static std::array<ExcelObj, _countof(theCellErrors)> theErrorObjs =
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

      static ExcelObj theEmptyString(PString::steal(L"\0"));
    }

    const ExcelObj& Missing()
    {
      return theMissing;
    }

    const ExcelObj& Error(CellError e)
    {
      switch (e)
      {
      case CellError::Null:        return theErrorObjs[0];
      case CellError::Div0:        return theErrorObjs[1];
      case CellError::Value:       return theErrorObjs[2];
      case CellError::Ref:         return theErrorObjs[3];
      case CellError::Name:        return theErrorObjs[4];
      case CellError::Num:         return theErrorObjs[5];
      case CellError::NA:          return theErrorObjs[6];
      case CellError::GettingData: return theErrorObjs[7];
      }
      XLO_THROW("Unexpected CellError type");
    }
    const ExcelObj& EmptyStr()
    {
      return theEmptyString;
    }
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
    case xltypeStr: return hash<wstring_view>()(value.cast<xloil::PStringRef>().view());
    case xltypeMissing:
    case xltypeNil: return 0;
    case xltypeErr: return hash<int>()(value.val.err);
    case xltypeSRef:
    {
      return xloil::boost_hash_combine(377,
        value.val.sref.ref.colFirst,
        value.val.sref.ref.colLast,
        value.val.sref.ref.rwFirst,
        value.val.sref.ref.rwLast);
    }
    case xltypeRef:
    {
      return xloil::boost_hash_combine(377, value.val.mref.idSheet, value.val.mref.lpmref);
    }
    case xltypeMulti:
      return hash<void*>()(value.val.array.lparray);
    default:
      return 0;
    }
  }
}
