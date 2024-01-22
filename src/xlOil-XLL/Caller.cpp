#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/State.h>
#include <xloil/ExcelArray.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/XlCallSlim.h>
#include <xlOilHelpers/Environment.h>
#include <array>

using namespace msxll;

namespace
{
  // Gross code to detect if being called by function wiz. Verbatim from:
  // https://docs.microsoft.com/en-us/office/client-developer/excel/how-to-call-xll-functions-from-the-function-wizard-or-replace-dialog-boxes?redirectedfrom=MSDN#Y241

  // Another possible way of detecting this is that MSO.dll will not be on
  // the call stack if called from a worksheet directly, although reading the 
  // call stack is not for free...

#define CLASS_NAME_BUFFSIZE  50
#define WINDOW_TEXT_BUFFSIZE  50
  // Data structure used as input to xldlg_enum_proc(), called by
  // called_from_paste_fn_dlg(), called_from_replace_dlg(), and
  // called_from_Excel_dlg(). These functions tell the caller whether
  // the current worksheet function was called from one or either of
  // these dialog boxes.
  struct xldlg_enum_struct
  {
    bool is_dlg;
    HWND  hwnd;
    const char *window_title_text; // set to NULL if don't care
    DWORD pid;
  };

#pragma warning(disable: 4311 4302)
  // The callback function called by Windows for every top-level window.
  bool CALLBACK xldlg_enum_proc(HWND hwnd, xldlg_enum_struct *p_enum)
  {
    // Check if the parent window is Excel.
    // Note: Because of the change from MDI (Excel 2010)
    // to SDI (Excel 2013) we check the process IDs
    if (p_enum->hwnd)
    {
      if (GetParent(hwnd) != p_enum->hwnd)
        return TRUE; // Tells Windows to continue iterating.
    }
    else
    {
      DWORD pid = NULL;
      GetWindowThreadProcessId(hwnd, &pid);
      if (pid != p_enum->pid)
        return TRUE;
    }
 
    char class_name[CLASS_NAME_BUFFSIZE + 1];
    //  Ensure that class_name is always null terminated for safety.
    class_name[CLASS_NAME_BUFFSIZE] = 0;
    GetClassNameA(hwnd, class_name, CLASS_NAME_BUFFSIZE);

    //  Do a case-insensitve comparison for the Excel dialog window
    //  class name with the Excel version number truncated.
    size_t len; // The length of the window's title text
    if (_strnicmp(class_name, "bosa_sdm_xl", 11) != 0)
      return TRUE;
    
    // Check if a searching for a specific title string
    if (p_enum->window_title_text)
    {
      // Get the window's title and see if it matches the given text.
      char buffer[WINDOW_TEXT_BUFFSIZE + 1];
      buffer[WINDOW_TEXT_BUFFSIZE] = 0;

      len = GetWindowTextA(hwnd, buffer, WINDOW_TEXT_BUFFSIZE);
      if (len == 0) // No title
      {
        if (p_enum->window_title_text[0] != 0)
          return TRUE; // No match, so keep iterating
      }
      // Window has a title so do a case-insensitive comparison of the
      // title and the search text, if provided.
      else if (p_enum->window_title_text[0] != 0
        && _stricmp(buffer, p_enum->window_title_text) != 0)
        return TRUE; // Keep iterating
    }
    p_enum->is_dlg = true;
    return FALSE; // Tells Windows to stop iterating.
  }

  bool called_from_paste_fn_dlg()
  {
    DWORD pid = 0;
    const char* windowName;
    auto& state = xloil::Environment::excelProcess();
    if (state.version < 13)
      windowName = "";
    else
    {
      windowName = "Function Arguments";
      pid = GetProcessId(GetCurrentProcess());
    }

    // Search for bosa_sdm_xl* dialog box with no title string.
    xldlg_enum_struct es = { false, (HWND)state.hWnd, windowName, pid };
    EnumWindows((WNDENUMPROC)xldlg_enum_proc, (LPARAM)&es);
    return es.is_dlg;
  }
}

namespace xloil
{
  namespace
  {
    // Max col is XFD, i.e 2^14, we run to XFZ
    constexpr size_t COL_NAME_CACHE_SIZE = XL_MAX_COLS + 22;
    constexpr size_t COL_NAME_WIDTH = 3;

    auto fillColumnNameCache()
    {
      std::array<char, COL_NAME_CACHE_SIZE * COL_NAME_WIDTH> result;
      auto* ptr = result.data();

      for (auto d = 'A'; d <= 'Z'; ++d)
      {
        *ptr++ = d;
        *ptr++ = 0;
        *ptr++ = 0;
      }

      for (auto c = 'A'; c <= 'Z'; ++c)
        for (auto d = 'A'; d <= 'Z'; ++d)
        {
          *ptr++ = c;
          *ptr++ = d;
          *ptr++ = 0;
        }

      for (auto c = 'A'; c <= 'X'; ++c)
        for (auto d = 'A'; d <= (c == 'X' ? 'F' : 'Z'); ++d)
          for (auto e = 'A'; e <= 'Z'; ++e)
          {
            *ptr++ = c;
            *ptr++ = d;
            *ptr++ = e;
          }

      assert(ptr == result.data() + result.size());

      return result;
    }

    static auto theColumnNameCache = fillColumnNameCache();

    template<class TChar>
    uint8_t writeColumn(size_t colIndex, TChar buf[3])
    {
      auto colName = &theColumnNameCache[colIndex * COL_NAME_WIDTH];
      buf[0] = colName[0];
      if (colIndex < 26)
        return 1;
      buf[1] = colName[1];
      if (colIndex < 26 + 26 * 26)
        return 2;
      buf[2] = colName[2];
      return 3;
    }

    uint8_t writeColumnNameW(size_t colIndex, wchar_t buf[3])
    {
      return writeColumn(colIndex, buf);
    }

    void writeDecimal(size_t value, wchar_t*& buf, size_t& bufSize)
    {
      auto nWritten = unsignedToString<10>(value, buf, bufSize);
      buf += nWritten;
      bufSize -= nWritten;
    }

    struct WriteA1
    {
      static constexpr size_t MAX_LEN = XL_CELL_ADDRESS_A1_MAX_LEN;
      void operator()(size_t row, size_t col, wchar_t*& buf, size_t& bufSize) const
      {
        auto nWritten = writeColumnNameW(col, buf);
        buf += nWritten;
        bufSize -= nWritten;

        writeDecimal(row + 1u, buf, bufSize);
      }
    };

    struct WriteRC
    {
      static constexpr size_t MAX_LEN = XL_CELL_ADDRESS_RC_MAX_LEN;
      void operator()(size_t row, size_t col, wchar_t*& buf, size_t& bufSize)
      {
        // Note we add one everywhere here as row/col is zero-based but 
        // A1/RC format is 1-based

        *buf++ = L'R';
        bufSize -= 1;
        writeDecimal(row + 1u, buf, bufSize);

        *buf++ = L'C';
        bufSize -= 1;
        writeDecimal(col + 1u, buf, bufSize);
      }
    };

    template<class TWriter>
    uint16_t writeLocalAddress(
      const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize)
    {
      // Rather than checking the bufSize at every step, just give up
      // if it can't hold the maxiumum possible size
      if (bufSize < TWriter::MAX_LEN)
        return 0;

      auto initialBuf = buf;
      TWriter writer;
      if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
      {
        // Single cell address
        writer(ref.rwFirst, ref.colFirst, buf, bufSize);
      }
      else
      {
        // Range address
        writer(ref.rwFirst, ref.colFirst, buf, bufSize);
        *buf++ = L':';
        --bufSize;
        writer(ref.rwLast, ref.colLast, buf, bufSize);
      }

      *buf = L'\0'; // Don't increment so we don't count terminator
      return (uint16_t)(buf - initialBuf);
    }

    uint16_t writeSheetAddress(
      wchar_t* buf,
      size_t bufLen,
      const msxll::XLREF12& sheetRef,
      const PStringRef& fullSheetName,
      const bool A1Style,
      const bool quoteSheetName)
    {
      uint16_t nWritten = 0;
      const auto sheetNamePStr = fullSheetName.pstr();
      const uint16_t sheetNameLength = fullSheetName.length();

      // There are some complicated but unwritten rules on when Excel quotes
      // the sheet name in an address.  See https://stackoverflow.com/questions/41677779/
      // We avoid the complexity of checking this and simply quote everything
      if (sheetNameLength > 0)
      {
        if (quoteSheetName)
        {
          if (bufLen <= sheetNameLength + 1u + 2)
            return 0;

          *(buf++) = L'\'';
          wmemcpy(buf, sheetNamePStr, sheetNameLength);
          buf += sheetNameLength;
          *(buf++) = L'\'';

          nWritten += 2;
        }
        else
        {
          if (bufLen <= sheetNameLength + 1u)
            return 0;
          wmemcpy(buf, sheetNamePStr, sheetNameLength);
          buf += sheetNameLength;
        }

        *(buf++) = L'!'; // Separator character
        nWritten += sheetNameLength + 1u;
        bufLen -= nWritten;
      }

      nWritten += A1Style
        ? writeLocalAddress<WriteA1>(sheetRef, buf, bufLen)
        : writeLocalAddress<WriteRC>(sheetRef, buf, bufLen);

      return nWritten;
    }

    int writeAddressImpl(
      wchar_t* buf,
      size_t bufLen,
      const ExcelObj& address,
      const PStringRef& sheetName,
      AddressStyle style)
    {
      const bool a1Style    = (style & AddressStyle::RC) == 0;
      const bool quoteSheet = (style & AddressStyle::NOQUOTE) == 0;
      switch (address.type())
      {
      case ExcelType::SRef:
      {
        return writeSheetAddress(buf, bufLen, address.val.sref.ref,
          sheetName, a1Style, quoteSheet);
      }
      case ExcelType::Ref:
      {
        return writeSheetAddress(buf, bufLen, address.val.mref.lpmref->reftbl[0],
          sheetName, a1Style, quoteSheet);
      }
      case ExcelType::Str: // Graphic object or Auto_Open/Auto_Close/Auto_Activate/... macro caller
      {
        // The sheet name is the active sheet at the time CallerInfo was created
        // Buttons can only be clicked on the active sheet I presume!
        const auto sheetNameLen = sheetName.length();
        const auto objectName = address.cast<PStringRef>();
        const uint16_t maxLen = 
          (a1Style ? XL_FULL_ADDRESS_A1_MAX_LEN : XL_FULL_ADDRESS_RC_MAX_LEN)
          + (quoteSheet ? 2 : 0) 
          - sheetNameLen;
        // Never return a string longer than the advertised max length
        const auto nameLen = std::min<uint16_t>(objectName.length(), maxLen);
        
        wmemcpy_s(
          buf, bufLen, 
          sheetName.pstr(), sheetNameLen);

        if (bufLen > sheetNameLen)
          wmemcpy_s(
            buf + sheetNameLen, bufLen - sheetNameLen, 
            objectName.pstr(), nameLen);

        return std::min<int>((int)bufLen, sheetNameLen + nameLen);
      }
      case ExcelType::Num: // DLL caller
      {
        return _snwprintf_s(buf, bufLen, bufLen, L"DLL(%d)", address.cast<int>());
      }
      case ExcelType::Multi:
      {
        ExcelArray arr(address, false);
        switch (arr.size())
        {
        case 2: // 2 element array describing a toolbar
          return _snwprintf_s(buf, bufLen, bufLen, L"Toolbar(%d)", arr.at(0).cast<int>());
        case 4: // 4 element array describing a menu
          return _snwprintf_s(buf, bufLen, bufLen, L"Menu(%d)", arr.at(0).cast<int>());
        default:
          XLO_THROW("Caller: address is badly formed array");
        }
      }
      default: // Other callers
        constexpr wchar_t nonWorksheetCaller[] = L"Unknown";
        if (bufLen < _countof(nonWorksheetCaller))
          return 0;
        wcscpy_s(buf, bufLen, nonWorksheetCaller);
        return _countof(nonWorksheetCaller);
      }
    }
  }

  namespace 
  {
    // A local reference used to get the active sheet name
    ExcelObj theA1Ref(msxll::xlref12{ 1, 1, 1, 1 });
  }

  CallerInfo::CallerInfo()
  {
    callExcelRaw(xlfCaller, &_address);
    callExcelRaw(xlSheetNm, &_sheetName, 
      _address.isType(ExcelType::RangeRef) ? &_address : &theA1Ref);
  }
  
  CallerInfo::CallerInfo(
    const ExcelObj& address, const wchar_t* fullSheetName)
    : _address(address)
  {
    if (fullSheetName)
      _sheetName = fullSheetName;
  }

  uint16_t CallerInfo::addressLength(AddressStyle style) const
  {
    // Any value in more precise guess?
    const auto sheetName = fullSheetName();
    if (!sheetName.empty())
    {
      const bool a1Style = (style & AddressStyle::RC) == 0;
      const bool quoteSheet = (style & AddressStyle::NOQUOTE) == 0;
      // Sheetname + quotes + ! + address
      return sheetName.length() + 1
        + a1Style ? XL_FULL_ADDRESS_A1_MAX_LEN : XL_FULL_ADDRESS_RC_MAX_LEN
        + quoteSheet ? 2 : 0;
    }

    // Not a worksheet caller
    auto addressStr = _address.cast<PStringRef>();
    if (!addressStr.empty())
      return addressStr.length();

    return 19; // Max is "Toolbar(<some int id>)"
  }

  int CallerInfo::writeAddress(
    wchar_t* buf, 
    size_t bufLen, 
    AddressStyle style) const
  {
    return writeAddressImpl(
      buf, bufLen, _address, _sheetName.cast<PStringRef>(), style);
  }
  
  std::wstring CallerInfo::address(AddressStyle style) const
  {
    std::wstring result;
    result.resize(addressLength(style));
    const auto nChars = writeAddress(result.data(), result.size(), style);
    result.resize(nChars);
    return result;
  }

  std::wstring CallerInfo::localAddress(AddressStyle style) const
  {
    wchar_t buf[XL_CELL_ADDRESS_RC_MAX_LEN];

    const msxll::XLREF12* sheetRef;
    switch (_address.type())
    {
    case ExcelType::SRef: sheetRef = &_address.val.sref.ref; break;
    case ExcelType::Ref:  sheetRef = _address.val.mref.lpmref->reftbl;  break;
    default:
      return std::wstring();
    }

    auto nWritten = style == AddressStyle::A1
      ? writeLocalAddress<WriteA1>(*sheetRef, buf, _countof(buf))
      : writeLocalAddress<WriteRC>(*sheetRef, buf, _countof(buf));

    return std::wstring(buf, nWritten);
  }

  uint8_t writeColumnName(size_t colIndex, char buf[4])
  {
    return writeColumn(colIndex, buf);
  }

  XLOIL_EXPORT uint16_t xlrefToLocalA1(
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize)
  {
    return writeLocalAddress<WriteA1>(ref, buf, bufSize);
  }

  XLOIL_EXPORT uint16_t xlrefToLocalRC(
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize)
  {
    return writeLocalAddress<WriteRC>(ref, buf, bufSize);
  }

  XLOIL_EXPORT uint16_t xlrefWriteWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style,
    bool quoteSheet)
  {
    ExcelObj sheetNm;
    sheetNm.xltype = msxll::xltypeRef;
    sheetNm.val.mref.idSheet = sheet;
    callExcelRaw(msxll::xlSheetNm, &sheetNm, &sheetNm);

    return writeSheetAddress(buf, bufSize, ref, sheetNm.cast<PStringRef>(), A1Style, quoteSheet);
  }

  namespace
  {
    struct ColumnNameAlphabet
    {
      static constexpr int8_t _alphabet[] = {
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
        -1, -1, -1, -1, -1, -1,
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
      };
      uint8_t operator()(int16_t c) const
      {
        return (uint8_t)((c < 'A' || c > 'z') ? -1 : _alphabet[c - 'A']);
      }
    };

    template<class Char, Char What>
    void skipOne(const Char*& c)
    {
      if (*c == What) ++c;
    }

    auto parseColLetters(const wchar_t*& c, const wchar_t* last)
    {
      // See notes above
      skipOne<wchar_t, L'$'>(c);
      // Look for a 3 char string as the column number.  A=1 and Z=26 but
      // there is no zero, so it is not base-27, hence we must use the 
      // `THigh` parameter of parseUnsigned.
      // 
      // We subtract 1 as A1-refs are 1-based but XLREF12 is zero based. 
      // This means failure to read anything will return -1, which gives an 
      // error condition later.
      auto end = std::min(c + 3, last);
      auto val = (int)detail::parseUnsigned<decltype(c), ColumnNameAlphabet, 26, 26>(c, end) - 1;
      return val > XL_MAX_COLS ? -1 : val;
    }
    /// <summary>
    /// Parse a number up to XL_MAX_COLS (16384), returning -1 on failure.
    /// </summary>
    template<size_t TNDigits, size_t TMaxVal>
    auto readNumberPart(const wchar_t*& c, const wchar_t* end)
    {
      // Skip the dollar symbol as it doesn't impact address conversion
      skipOne<wchar_t, L'$'>(c);
      end = std::min(c + TNDigits, end);
      int val = (int)parseUnsigned<10>(c, end) - 1;
      return val > TMaxVal ? -1 : val;
    }

    template<wchar_t TPrefix, size_t TNDigits, size_t TMaxVal>
    auto readPrefixedNumber(const wchar_t*& c, const wchar_t* end)
    {
      if (c == end)
        return -1;
      skipOne<wchar_t, L'$'>(c);
      if (*c == TPrefix || *c == TPrefix + (L'a' - L'A'))
      {
        ++c;
        return readNumberPart<TNDigits, TMaxVal>(c, end);
      }
      return -1;
    }

    constexpr auto(*parseR)(const wchar_t*&, const wchar_t*) = &readPrefixedNumber<L'R', 7, XL_MAX_ROWS>;
    constexpr auto(*parseC)(const wchar_t*&, const wchar_t*) = &readPrefixedNumber<L'C', 5, XL_MAX_COLS>;
    constexpr auto(*parseRowNumber)(const wchar_t*&, const wchar_t*) = &readNumberPart<7, XL_MAX_ROWS>;
  }
  
  bool localAddressToXlRef(msxll::XLREF12& r, const std::wstring_view& address)
  {
    const wchar_t* c = address.data();
    const wchar_t* end = c + address.size();
    memset(&r, 0, sizeof(decltype(r)));

    r.colFirst = parseColLetters(c, end);
    r.rwFirst = parseRowNumber(c, end);
    
    if (c == end)
    {
      r.colLast = r.colFirst;
      r.rwLast = r.rwFirst;
    } 
    else if (*c == L':') // Look for the address separator
    {
      ++c;
      r.colLast = parseColLetters(c, end);
      r.rwLast = parseRowNumber(c, end);
    }
    else 
    {
      // Failed to read address as A1-type, try RC-type.  We know that 
      // The first Rxxx will have been read by the A1-parser as the 
      // row number.
      r.colFirst = parseC(c, end);
      if (c == end)
      {
        r.colLast = r.colFirst;
        r.rwLast = r.rwFirst;
      }
      else if (*c == L':')
      {
        ++c;
        r.rwLast = parseR(c, end);
        r.colLast = parseC(c, end);
      }
      else
        return false;
    }

    // Return true if parsing was successful
    return r.colFirst >= 0 && r.rwFirst >= 0 && r.rwLast >= 0 && r.colLast >= 0;
  }

  bool inFunctionWizard()
  {
    return called_from_paste_fn_dlg();
  }
}
