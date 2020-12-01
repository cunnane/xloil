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

      // TODO: does this need adapting for other locales?
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
    auto& state = xloil::State::excelState();
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
  CallerInfo::CallerInfo()
  {
    callExcelRaw(xlfCaller, &_address);
    if (_address.isType(ExcelType::RangeRef))
      callExcelRaw(xlSheetNm, &_fullSheetName, &_address);
  }
  CallerInfo::CallerInfo(
    const ExcelObj& address, const wchar_t* fullSheetName)
    : _address(address)
  {
    if (fullSheetName)
      _fullSheetName = fullSheetName;
  }
  uint16_t CallerInfo::addressRCLength() const
  {
    // Any value in more precise guess?
    const auto sheetName = fullSheetName();
    if (!sheetName.empty())
      return sheetName.length() + 1 + XL_CELL_ADDRESS_RC_MAX_LEN;

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
      const auto wsLength = sheetName.length();
      if (wsLength > 0)
      {
        if (bufLen <= wsLength + 1)
          return 0;
        wmemcpy(buf, wsName, wsLength);
        buf += wsLength;
        // Separator character
        *(buf++) = L'!';

        nWritten += wsLength + 1;
        bufLen -= nWritten;
      }

      nWritten += A1Style
        ? xlrefToLocalA1(sheetRef, buf, bufLen)
        : xlrefToLocalRC(sheetRef, buf, bufLen);

      return nWritten;
    }
  }
  int CallerInfo::writeAddress(
    wchar_t* buf, 
    size_t bufLen,
    bool A1Style) const
  {
    switch (_address.type())
    {
    case ExcelType::SRef:
    {
      return writeSheetRef(buf, bufLen, _address.val.sref.ref,
        _fullSheetName.asPString(), A1Style);
    }
    case ExcelType::Ref:
    {
      return writeSheetRef(buf, bufLen, _address.val.mref.lpmref->reftbl[0],
        _fullSheetName.asPString(), A1Style);
    }
    case ExcelType::Str: // Graphic object or Auto_XXX macro caller
    {
      auto str = _address.asPString();
      wcsncpy_s(buf, bufLen, str.pstr(), str.length());
      return std::min<int>((int)bufLen, str.length());
    }
    case ExcelType::Num: // DLL caller
    {
      return _snwprintf_s(buf, bufLen, bufLen, L"DLL(%d)", _address.asInt());
    }
    case ExcelType::Multi:
    {
      ExcelArray arr(_address, false);
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

  std::wstring CallerInfo::writeAddress(bool A1Style) const
  {
    std::wstring result;
    result.resize(addressRCLength()); // RC-length > A1-length
    const auto nChars = writeAddress(result.data(), result.size(), A1Style);
    result.resize(nChars);
    return result;
  }

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

  void writeColumnName(size_t colIndex, char buf[4])
  {
    if (colIndex < COL_NAME_CACHE_SIZE)
    {
      memcpy_s(buf, 4, &theColumnNameCache[colIndex * 2], 2);
      buf[2] = '\0';
    }
    else
    {
      constexpr short Ato0 = 'A' - '0';
      constexpr short Atoa = 'A' - 'a' + 10;

      _itoa_s((int)colIndex - 26, buf, 4, 26);
      buf[0] += (buf[0] < 'A' ? Ato0 : Atoa) - 1;
      buf[1] += buf[1] < 'A' ? Ato0 : Atoa;
      buf[2] += buf[2] < 'A' ? Ato0 : Atoa;
    }
  }

  void writeColumnNameW(size_t colIndex, wchar_t buf[4])
  {
    size_t dummy;
    char colBuf[4];
    writeColumnName(colIndex, colBuf);
    mbstowcs_s(&dummy, buf, 4, colBuf, 4);
  }

  XLOIL_EXPORT uint16_t xlrefToLocalA1(
    const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize)
  {
    int ret;
    // Add one everywhere here as rwFirst is zero-based but A1 format is 1-based
    if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
    {
      wchar_t wcol[4];
      writeColumnNameW(ref.colFirst, wcol);
      ret = _snwprintf_s(buf, bufSize, bufSize, L"%s%d", wcol, ref.rwFirst + 1);
    }
    else
    {
      wchar_t wcolFirst[4], wcolLast[4];
      writeColumnNameW(ref.colFirst, wcolFirst);
      writeColumnNameW(ref.colLast, wcolLast);
      ret = _snwprintf_s(buf, bufSize, bufSize, L"%s%d:%s%d",
        wcolFirst, ref.rwFirst + 1, wcolLast, ref.rwLast + 1);
    }
    return ret < 0 ? 0 : (uint16_t)ret;
  }

  XLOIL_EXPORT uint16_t xlrefSheetAddress(
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

  XLOIL_EXPORT std::wstring xlrefSheetAddress(
    const msxll::IDSHEET& sheet,
    const msxll::XLREF12& ref,
    bool A1Style)
  {
    return captureWStringBuffer([&](auto buf, auto sz)
    {
      return xlrefSheetAddress(sheet, ref, buf, sz, A1Style);
    });
  }

  XLOIL_EXPORT std::wstring xlrefLocalAddress(
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

  bool inFunctionWizard()
  {
    return called_from_paste_fn_dlg();
  }
}
