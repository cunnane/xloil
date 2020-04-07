#include "xloil/Options.h"
#include "xloil/WindowsSlim.h"
#include "XlCallSlim.h"
#include "xloil/WindowsSlim.h"
#include "xloil/EntryPoint.h"
#include "xloil/ExportMacro.h"
#include "xloil/Log.h"
#include "xloil/Events.h"
#include "xloil/ExcelCall.h"
#include <delayimp.h>
#include <filesystem>
namespace fs = std::filesystem;

using std::wstring;

namespace
{
  static wstring ourXllPath;
  static wstring ourXllDir;

  bool setDllPath(HMODULE handle)
  {
    try
    {
      wchar_t path[4 * MAX_PATH]; // TODO: may not be long enough!!!
      auto size = GetModuleFileName(handle, path, sizeof(path));
      if (size == 0)
      {
        OutputDebugStringW(L"xloil_Loader: Could not determine XLL location");
        //XLO_ERROR(L"Could not determine XLL location: {}", writeWindowsError());
        return false;
      }
      ourXllPath = wstring(path, path + size);
      ourXllDir = fs::path(ourXllPath).remove_filename().wstring();
      return true;
    }
    catch (...)
    {
      return false;
    }
  }
}

namespace
{
  static HMODULE theModuleHandle = nullptr;

  FARPROC WINAPI coreLoaderHook(unsigned dliNotify, PDelayLoadInfo pdli)
  {
    // TODO: can we auto-set the hard-coded string here?
    switch (dliNotify)
    {
    case dliNotePreLoadLibrary:
      if (_stricmp(pdli->szDll, "xlOil_Loader") == 0)
        return (FARPROC)theModuleHandle;
      break;
    default:
      return NULL;
    }

    return NULL;
  }
}

XLO_ENTRY_POINT(int) DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    lpvReserved
)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
  {
    theModuleHandle = hinstDLL;
    if (!setDllPath(hinstDLL))
      return FALSE;
    xloil::coreInit(&coreLoaderHook, ourXllPath.c_str());
  }
  return TRUE;
}

/*
** xlAutoOpen
**
** xlAutoOpen is how Microsoft Excel loads XLL files.
** When you open an XLL, Microsoft Excel calls the xlAutoOpen
** function, and nothing more.
**/

XLO_ENTRY_POINT(int) xlAutoOpen(void)
{
  SetDllDirectory(ourXllDir.c_str());
  //XLO_TRACE("xlAutoOpen called in Loader");
  xloil::coreAutoOpen();
  xloil::Event_AutoOpen().fire();

  // TODO: handle failure?
  xloil::tryCallExcel(msxll::xlEventRegister, "xloHandleCalculationEnded", xleventCalculationEnded);
  xloil::tryCallExcel(msxll::xlEventRegister, "xloHandleCalculationCancelled", xleventCalculationCanceled);

  SetDllDirectory(NULL);
  return 1;
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  //XLO_TRACE("xlAutoClose called in Loader");
  xloil::coreAutoClose();
  return 1;
}

// This function can be called without the add-in loaded, so avoid using
// any xlOil functionality
XLO_ENTRY_POINT(msxll::xloper12*) xlAddInManagerInfo12(msxll::xloper12* xAction)
{
  int action = 0;
  switch (xAction->xltype)
  {
  case msxll::xltypeNum:
    action = xAction->val.num;
    break;
  case msxll::xltypeInt:
    action = xAction->val.w;
    break;
  }
  
  static msxll::xloper12 xInfo;
  if (action == 1)
  {
    xInfo.xltype = msxll::xltypeStr;
    xInfo.val.str = L"\011xlOil Addin";
  }
  else
  {
    xInfo.xltype = msxll::xltypeErr;
    xInfo.val.err = msxll::xlerrValue;
  }
  
  return &xInfo;
}

XLO_ENTRY_POINT(int) xlAutoFree12(msxll::xloper12* pxFree)
{
  delete (xloil::ExcelObj*)pxFree;
  return 1;
}

XLO_ENTRY_POINT(int) xloHandleCalculationEnded()
{
  xloil::Event_CalcEnded().fire();
  return 1;
}

XLO_ENTRY_POINT(int) xloHandleCalculationCancelled()
{
  xloil::Event_CalcCancelled().fire();
  return 1;
}