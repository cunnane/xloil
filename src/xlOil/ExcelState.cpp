#include "ExcelState.h"
#include "ExcelCall.h"
#include "EntryPoint.h"
#include <xlOilHelpers/WindowsSlim.h>


using namespace msxll;

namespace
{

  // Gross code to detect if being called by function wiz. Verbatim from:
  // https://docs.microsoft.com/en-us/office/client-developer/excel/how-to-call-xll-functions-from-the-function-wizard-or-replace-dialog-boxes?redirectedfrom=MSDN#Y241

  // Another possible way of detecting this is that MSO.dll will not be on the call stack if called from a worksheet directly.

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
    short low_hwnd;
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
    if (p_enum->low_hwnd)
    {
      if (LOWORD((DWORD)GetParent(hwnd)) != p_enum->low_hwnd)
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
    short hwnd = 0;
    DWORD pid = 0;
    const char* windowName;
    if (xloil::coreExcelVersion() < 13)
    {
      XLOPER xHwnd;
      // Calls Excel4, which only returns the low part of the Excel
      // main window handle. This is OK for the search however.
      if (Excel4(xlGetHwnd, &xHwnd, 0))
        return false; // Couldn't get it, so assume not
                      
      hwnd = xHwnd.val.w;
      windowName = "";
    }
    else
    {
      windowName = "Function Arguments";
      pid = GetProcessId(GetCurrentProcess());
    }

    // Search for bosa_sdm_xl* dialog box with no title string.
    xldlg_enum_struct es = { FALSE, hwnd, windowName, pid };
    EnumWindows((WNDENUMPROC)xldlg_enum_proc, (LPARAM)&es);
    return es.is_dlg;
  }
}

namespace xloil
{

  CallerInfo::CallerInfo()
    : _Address(new ExcelObj())
    , _SheetName(new ExcelObj())
  {
    callExcelRaw(xlfCaller, const_cast<ExcelObj*>(_Address.get()));
    if (_Address->isType(ExcelType::RangeRef))
      callExcelRaw(xlSheetNm, const_cast<ExcelObj*>(_SheetName.get()), _Address.get());
  }
  size_t CallerInfo::fullAddressLength() const
  {
    auto s = _SheetName->asPascalStr();
    // Any value in more precise guess?
    return s.length() + 1 + CELL_ADDRESS_RC_MAX_LEN;
  }
  size_t CallerInfo::writeFullAddress(wchar_t* buf, size_t bufLen) const
  {
    const auto wsName = _SheetName->asPascalStr();
    const auto wsLength = wsName.length();

    if (wsLength > 0)
    {
      assert(bufLen > wsLength);
      wmemcpy(buf, wsName.pstr(), wsLength);
      buf += wsLength;
      // Separator character
      *(buf++) = L'!';
    }

    // TODO: handle other caller cases?
    size_t addressLen;
    switch (_Address->type())
    {
    case ExcelType::SRef:
      addressLen = xlrefToStringRC(_Address->val.sref.ref, buf, bufLen - wsName.length() - 1);
      break;
    case ExcelType::Ref:
      addressLen = xlrefToStringRC(_Address->val.mref.lpmref->reftbl[0], buf, bufLen - wsName.length() - 1);
      break;
    default:
      wmemcpy(buf, L"UNKNOWN", 7);
      addressLen = 7;
      break;
    }
    return addressLen + wsLength + 1;
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

  XLOIL_EXPORT size_t xlrefToStringA1(const msxll::XLREF12& ref, wchar_t* buf, size_t bufSize)
  {
    // Add one everywhere here as rwFirst is zero-based but A1 format is 1-based
    if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
    {
      wchar_t wcol[4];
      writeColumnNameW(ref.colFirst, wcol);
      return _snwprintf_s(buf, bufSize, bufSize, L"%s%d", wcol, ref.rwFirst + 1);
    }
    else
    {
      wchar_t wcolFirst[4], wcolLast[4];
      writeColumnNameW(ref.colFirst, wcolFirst);
      writeColumnNameW(ref.colLast, wcolLast);
      return _snwprintf_s(buf, bufSize, bufSize, L"%s%d:%s%d", 
        wcolFirst, ref.rwFirst + 1, wcolLast, ref.rwLast + 1);
    }
  }

  XLOIL_EXPORT size_t xlrefSheetAddressA1(
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

    auto wsName = sheetNm.asPascalStr();
    assert(bufSize > wsName.length());
    wmemcpy(buf, wsName.pstr(), wsName.length());
    buf += wsName.length();

    // Separator character
    *(buf++) = L'!';
      
    auto addressLen = A1Style 
      ? xlrefToStringA1(ref, buf, bufSize - wsName.length() - 1)
      : xlrefToStringRC(ref, buf, bufSize - wsName.length() - 1);

    return addressLen + wsName.length() + 1;
  }

  // Uses RxCy format as it's easier for the programmer 
  // (see how much code is required above for A1 style!)
  size_t xlrefToStringRC(const XLREF12& ref, wchar_t* buf, size_t bufSize)
  {
    // Add one everywhere here as rwFirst is zero-based but RxCy format is 1-based
    if (ref.rwFirst == ref.rwLast && ref.colFirst == ref.colLast)
      return _snwprintf_s(buf, bufSize, bufSize, L"R%dC%d", ref.rwFirst + 1, ref.colFirst + 1);
    else
      return _snwprintf_s(buf, bufSize, bufSize, L"R%dC%d:R%dC%d", 
        ref.rwFirst + 1, ref.colFirst + 1, ref.rwLast + 1, ref.colLast + 1);
  }

  bool inFunctionWizard()
  {
    return called_from_paste_fn_dlg();
  }
}
