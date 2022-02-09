#include <xloil/Caller.h>
#include <xloil/ExcelCall.h>
#include <xloil/State.h>
#include <xloil/ExcelArray.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/XlCallSlim.h>
#include <xloil/AppObjects.h>
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
    auto& state = xloil::App::internals();
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
    uint16_t writeSheetRef(
      wchar_t* buf,
      size_t bufLen,
      const msxll::XLREF12& sheetRef,
      const PString<>& sheetName,
      const bool A1Style)
    {
      uint16_t nWritten = 0;
      const auto wsName = sheetName.pstr();
      const uint16_t wsLength = sheetName.length();
      if (wsLength > 0)
      {
        if (bufLen <= wsLength + 1u)
          return 0;
        wmemcpy(buf, wsName, wsLength);
        buf += wsLength;
        // Separator character
        *(buf++) = L'!';

        nWritten += wsLength + 1u;
        bufLen -= nWritten;
      }

      nWritten += A1Style
        ? xlrefToLocalA1(sheetRef, buf, bufLen)
        : xlrefToLocalRC(sheetRef, buf, bufLen);

      return nWritten;
    }

    uint16_t writeInternal(
      wchar_t* buf,
      size_t bufLen,
      const msxll::XLREF12& sheetRef,
      const msxll::IDSHEET sheetId)
    {
      const auto cellStart = sheetRef.rwFirst * XL_MAX_COLS + sheetRef.colFirst;
      const auto cellEnd   = sheetRef.rwLast * XL_MAX_COLS + sheetRef.colLast;
      const auto nWritten  = _snwprintf_s(buf, bufLen, bufLen, L"[%p]%x:%x", sheetId, cellStart, cellEnd);
      return nWritten < 0 ? 0 : (uint16_t)nWritten;
    }

    int writeAddressImpl(
      wchar_t* buf,
      size_t bufLen,
      const ExcelObj& address,
      const PString<>& sheetName,
      CallerInfo::AddressStyle style)
    {
      switch (address.type())
      {
      case ExcelType::SRef:
      {
        switch (style)
        {
        case CallerInfo::RC:
        case CallerInfo::A1:
          return writeSheetRef(buf, bufLen, address.val.sref.ref,
            sheetName, style == CallerInfo::A1);
        }
      }
      case ExcelType::Ref:
      {
        switch (style)
        {
        case CallerInfo::RC:
        case CallerInfo::A1:
          return writeSheetRef(buf, bufLen, address.val.mref.lpmref->reftbl[0],
            sheetName, style == CallerInfo::A1);
        }
      }
      case ExcelType::Str: // Graphic object or Auto_XXX macro caller
      {
        auto str = address.asPString();
        // Never return a string longer than the advertised max length
        uint16_t maxLen = str.length();
        switch (style)
        {
        case CallerInfo::RC:       maxLen = std::min(maxLen, XL_FULL_ADDRESS_RC_MAX_LEN); break;
        case CallerInfo::A1:       maxLen = std::min(maxLen, XL_FULL_ADDRESS_A1_MAX_LEN); break;
        }
        wmemcpy_s(buf, bufLen, str.pstr(), maxLen);
        return std::min<int>((int)bufLen, maxLen);
      }
      case ExcelType::Num: // DLL caller
      {
        return _snwprintf_s(buf, bufLen, bufLen, L"DLL(%d)", address.asInt());
      }
      case ExcelType::Multi:
      {
        ExcelArray arr(address, false);
        switch (arr.size())
        {
        case 2:
          return _snwprintf_s(buf, bufLen, bufLen, L"Toolbar(%d)", arr.at(0).asInt());
        case 4:
          return _snwprintf_s(buf, bufLen, bufLen, L"Menu(%d)", arr.at(0).asInt());
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

  CallerInfo::CallerInfo()
  {
    callExcelRaw(xlfCaller, &_address);
    if (_address.isType(ExcelType::RangeRef))
      callExcelRaw(xlSheetNm, &_sheetName, &_address);
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
      switch (style)
      {
        case RC: return sheetName.length() + 1 + XL_CELL_ADDRESS_RC_MAX_LEN;
        case A1: return sheetName.length() + 1 + XL_CELL_ADDRESS_A1_MAX_LEN;
      }
    }

    // Not a worksheet caller
    auto addressStr = _address.asPString();
    if (!addressStr.empty())
      return addressStr.length();

    return 19; // Max is "Toolbar(<some int id>)"
  }

  namespace
  {
    template<size_t N>
    constexpr size_t wcslength(wchar_t const (&)[N])
    {
      return N - 1;
    }
  }
  int CallerInfo::writeAddress(wchar_t* buf, size_t bufLen, AddressStyle style) const
  {
    return writeAddressImpl(buf, bufLen, _address, _sheetName.asPString(), style);
  }
  std::wstring CallerInfo::writeAddress(AddressStyle style) const
  {
    std::wstring result;
    result.resize(addressLength(style));
    const auto nChars = writeAddress(result.data(), result.size(), style);
    result.resize(nChars);
    return result;
  }

  namespace
  {
    constexpr size_t COL_NAME_CACHE_SIZE = 26 + 26 * 26;

    auto fillColumnNameCache()
    {
      static std::array<char, COL_NAME_CACHE_SIZE * 2> cache;
      auto* pcolumns = cache.data();

      for (auto d = 'A'; d <= 'Z'; ++d, pcolumns += 2)
      {
        pcolumns[0] = d;
        pcolumns[1] = 0;
      }

      for (auto c = 'A'; c <= 'Z'; ++c)
        for (auto d = 'A'; d <= 'Z'; ++d, pcolumns += 2)
        {
          pcolumns[0] = c;
          pcolumns[1] = d;
        }
      return cache;
    }

    static auto theColumnNameCache = fillColumnNameCache();
  }

  uint8_t writeColumnName(size_t colIndex, char buf[4])
  {
    // We cache everything from A to ZZ
    if (colIndex < COL_NAME_CACHE_SIZE)
    {
      auto colName = &theColumnNameCache[colIndex * 2];
      buf[0] = colName[0];
      if (colIndex < 26)
        return 1;
      buf[1] = colName[1];
      return 2;
    }
    else
    {
      // If we get here we must have a 3 letter column name

      constexpr short Ato0 = 'A' - '0';
      constexpr short Atoa = 'A' - 'a' + 10;

      // We write the number base-26, then shift the characters to bring
      // them into the 'A-Z' rangee
      _itoa_s((int)colIndex - 26, buf, 4, 26);
      buf[0] += (buf[0] < 'A' ? Ato0 : Atoa) - 1;
      buf[1] += buf[1] < 'A' ? Ato0 : Atoa;
      buf[2] += buf[2] < 'A' ? Ato0 : Atoa;

      return 3;
    }
  }

  uint8_t writeColumnNameW(size_t colIndex, wchar_t buf[3])
  {
    if (colIndex < COL_NAME_CACHE_SIZE)
    {
      auto colName = &theColumnNameCache[colIndex * 2];
      buf[0] = colName[0];
      if (colIndex < 26)
        return 1;
      buf[1] = colName[1];
      return 2;
    }
    else
    {
      char colName[4];
      writeColumnName(colIndex, colName);
      buf[0] = colName[0];
      buf[1] = colName[1];
      buf[2] = colName[2];
      return 3;
    }
  }

  XLOIL_EXPORT uint16_t xlrefToLocalA1(
    const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize)
  {
    // Rather than checking the bufSize at every step, just give up
    // if it can't hold the maxiumum possible size
    if (bufSize < XL_CELL_ADDRESS_A1_MAX_LEN)
      return 0;

    uint16_t nWritten;
    // Note we add one everywhere here as rwFirst is zero-based but A1 format is 1-based
    if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
    {
      // Single cell address
      nWritten = writeColumnNameW(ref.colFirst, buf);
      nWritten += unsignedToString<10>((unsigned)ref.rwFirst + 1, buf + nWritten, bufSize - nWritten);
    }
    else
    {
      // Range address
      nWritten = writeColumnNameW(ref.colFirst, buf);
      nWritten += unsignedToString<10>((unsigned)ref.rwFirst + 1, buf + nWritten, bufSize - nWritten);

      buf[nWritten] = L':';
      ++nWritten;

      nWritten += writeColumnNameW(ref.colLast, buf + nWritten);
      nWritten += unsignedToString<10>((unsigned)ref.rwLast + 1, buf + nWritten, bufSize - nWritten);
    }

    buf[nWritten] = L'\0';
    return nWritten;
  }

  XLOIL_EXPORT uint16_t xlrefWriteWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    wchar_t* buf,
    size_t bufSize,
    bool A1Style)
  {
    ExcelObj sheetNm;
    sheetNm.xltype = msxll::xltypeRef;
    sheetNm.val.mref.idSheet = sheet;
    callExcelRaw(msxll::xlSheetNm, &sheetNm, &sheetNm);

    return writeSheetRef(buf, bufSize, ref, sheetNm.asPString(), A1Style);
  }

  XLOIL_EXPORT std::wstring xlrefToWorkbookAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    bool A1Style)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
    {
      return xlrefWriteWorkbookAddress(sheet, ref, buf, sz, A1Style);
    });
  }

  XLOIL_EXPORT std::wstring xlrefToLocalAddress(
    const msxll::XLREF12& ref,
    bool A1Style)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
      {
        return A1Style
          ? xlrefToLocalA1(ref, buf, sz)
          : xlrefToLocalRC(ref, buf, sz);
      },
      XL_CELL_ADDRESS_A1_MAX_LEN);
  }

  // Uses RxCy format as it's easier for the programmer 
  // (see how much code is required above for A1 style!)
  uint16_t xlrefToLocalRC(const XLREF12& ref, wchar_t* buf, size_t bufSize)
  {
    int ret;
    // Add one everywhere here as rwFirst is zero-based but RxCy format is 1-based
    if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
      ret = _snwprintf_s(buf, bufSize, bufSize, L"R%dC%d", ref.rwFirst + 1, ref.colFirst + 1);
    else
      ret = _snwprintf_s(buf, bufSize, bufSize, L"R%dC%d:R%dC%d",
        ref.rwFirst + 1, ref.colFirst + 1, ref.rwLast + 1, ref.colLast + 1);
    return ret < 0 ? 0 : (uint16_t)ret;
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
      uint8_t operator()(int8_t c) const
      {
        return (uint8_t)((c < 'A' || c > 'z') ? -1 : _alphabet[c - 'A']);
      }
    };

    template<class Char, Char What>
    void skipOne(const Char*& c)
    {
      if (*c == What) ++c;
    }

    auto parseRow(const wchar_t*& c, const wchar_t* end)
    {
      // Skip the dollar symbol as it doesn't impact address conversion
      skipOne<wchar_t, L'$'>(c);
      // Look for a base-10 number as the row number. We only read 7 digits so
      // there is no chance of overflow. We subtract 1 as A1-refs are 1-based 
      // but XLREF12 is zero based. This means failure to read anything will 
      // return -1, which gives an error condition later.
      end = std::min(c + 7, end);
      int val = (int)parseUnsigned<10>(c, end) - 1; 
      return val > XL_MAX_ROWS ? -1 : val;
    }
    auto parseCol(const wchar_t*& c, const wchar_t* last)
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
  }
  
  bool localAddressToXlRef(msxll::XLREF12& r, const std::wstring_view& address)
  {
    const wchar_t* c = address.data();
    const wchar_t* end = c + address.size();
    memset(&r, 0, sizeof(decltype(r)));

    r.colFirst = parseCol(c, end);
    r.rwFirst = parseRow(c, end);

    // Look for the address separator
    if (c < end && *c == L':')
    {
      ++c;
      r.colLast = parseCol(c, end);
      r.rwLast = parseRow(c, end);
    }
    else if (c == end) 
    {
      r.colLast = r.colFirst;
      r.rwLast  = r.rwFirst;
    }
    else
      return false; // Trailing unparsable characters

    // Return true if parsing was successful
    return r.colFirst >= 0 && r.rwFirst >= 0 && r.rwLast >= 0 && r.colLast >= 0;
  }

  bool inFunctionWizard()
  {
    return called_from_paste_fn_dlg();
  }
}
