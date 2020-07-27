#include <xloilHelpers/WindowsSlim.h>
#include <xloilHelpers/Environment.h>
#include <xloilHelpers/Settings.h>
#include <xlOil/XlCallSlim.h>
#include <xloil/Loaders/EntryPoint.h>
#include <xloil/ExportMacro.h>
#include <xloil/ExcelCall.h>
#include <tomlplusplus/toml.hpp>
#include <filesystem>
#include <delayimp.h>
#include <fstream>

namespace fs = std::filesystem;

using std::string;
using std::wstring;
using namespace xloil;
using std::vector;
using std::shared_ptr;
using namespace std::string_literals;

namespace
{
  static wstring ourXllPath;
  // This bool is required due to apparent bugs in the XLL interface:
  // Excel may call XLL event handlers after calling xlAutoClose,
  // and it may call xlAutoClose without ever having called xlAutoOpen
  // This former to happen when Excel is closing and asks the user 
  // to save the workbook, the latter when removing an addin using COM
  // automation
  static bool theXllIsOpen = false;

  bool setDllPath(HMODULE handle)
  {
    try
    {
      wchar_t path[4 * MAX_PATH]; // TODO: may not be long enough!!!
      auto size = GetModuleFileName(handle, path, sizeof(path));
      if (size == 0)
      {
        OutputDebugStringW(L"xloil_Loader: Could not determine XLL location");
        return false;
      }
      ourXllPath = wstring(path, path + size);
      return true;
    }
    catch (...)
    {
      return false;
    }
  }

  std::unique_ptr<std::fstream> theLogFile;

  /// <summary>
  /// Very cheap log file to catch startup errors before
  /// the core dll can initialise spdlog.
  /// </summary>
  void logError(const char* err)
  {
    OutputDebugStringA(err);
    if (!theLogFile)
      theLogFile.reset(new std::fstream(fs::path(ourXllPath).replace_extension("log"), std::ios::app));
    *theLogFile << err << "\n";
    theLogFile->flush();
  }

  // Avoids using xloil so we can call before the dll is found
  int getExcelVersion()
  {
    using namespace msxll;

    // https://github.com/MicrosoftDocs/office-developer-client-docs/blob/...
    // master/docs/excel/calling-into-excel-from-the-dll-or-xll.md
    XLOPER arg, result;
    arg.xltype = xltypeInt;
    arg.val.w = 2;

    auto ret = Excel4(xlfGetWorkspace, &result, 1, &arg);
    if (ret != xlretSuccess || result.xltype != xltypeStr)
      return 0;
    auto pStr = result.val.str;
    auto versionStr = std::string(pStr + 1, pStr + 1 + pStr[0]);
    return std::stoi(versionStr);
  }
}

/// <summary>
/// Hook for delay load failures so we an return a sensible error
/// if xlOil.dll is not found
/// </summary>
FARPROC WINAPI delayLoadFailureHook(unsigned dliNotify, DelayLoadInfo * pdli)
{
  std::string msg;
  switch (dliNotify)
  {
  case dliFailGetProc:
    msg = formatStr("Unable to find procedure: %s in %s", pdli->dlp.szProcName, pdli->szDll);
  default:
    msg = formatStr("Unable to load library: %s", pdli->szDll);
  }
  logError(msg.c_str());
  return nullptr;
}

extern "C" const PfnDliHook __pfnDliFailureHook2 = delayLoadFailureHook;

namespace
{
  static HMODULE theModuleHandle = nullptr;
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
  }
  return TRUE;
}


// Name of core dll. Not sure if there's an automatic way to get this
constexpr wchar_t* const xloil_dll = L"XLOIL.DLL";

void loadEnvironmentBlock(const toml::table& settings)
{
  auto environment = Settings::environmentVariables(settings["Addin"]);

  for (auto&[key, val] : environment)
  {
    auto value = expandWindowsRegistryStrings(
      expandEnvironmentStrings(val));

    SetEnvironmentVariable(key.c_str(), value.c_str());
  }
}
int loadCore(const wchar_t* ourXllPath)
{
  // If the delay load fails, it will throw a SEH exception, so we must use
  // __try/__except to avoid this crashing Excel.
  int ret = -1;
  __try
  {
    ret = xloil::coreAutoOpen(ourXllPath);
  }
  __except (EXCEPTION_EXECUTE_HANDLER)
  {
    // TODO: add GetExceptionCode(), without using string
    logError("Failed to load xlOil.dll");
  }
  return ret;
}

/// <summary>
/// xlAutoOpen is how Microsoft Excel loads XLL files.
/// When you open an XLL, Microsoft Excel calls the xlAutoOpen
/// function, and nothing more.
/// </summary>
/// <returns>Must return 1</returns>
XLO_ENTRY_POINT(int) xlAutoOpen(void)
{
  try
  {
    // We need to find xloil.dll. 
    const fs::path ourXllDir = fs::path(ourXllPath).remove_filename();

    if (GetModuleHandle(xloil_dll) != 0) // Is it already loaded?
    {
      // Nothing to do
    }
    else if (fs::exists(ourXllDir / xloil_dll)) // Check same directory as XLL
    {
      OutputDebugStringW((wstring(L"xlOil Loader: SetDllDirectory = ") + ourXllDir.wstring()).c_str());
      SetDllDirectory(ourXllDir.c_str());
    }
    else
    {
      // Load as many environment blocks as we can and hope the dynamic loader
      // finds the DLL!  We used to look through the installed addins for 
      // xloil.xll, but since we now load this using the XLSTART folder, 
      // it won't appear in the registry

      // Load the environment block from our settings file (if it exists)
      auto settings = findSettingsFile(ourXllPath.c_str());
      if (settings)
        loadEnvironmentBlock(*settings);

      // Look for xloil.ini in our directory or AppData
      if (_wcsicmp(fs::path(ourXllPath).filename().c_str(), L"xloil.xll") != 0)
      {
        auto coreSettings = findSettingsFile(
          fs::path(ourXllPath).replace_filename(xloil_dll).c_str());
        if (coreSettings)
          loadEnvironmentBlock(*settings);
      }
    }

    auto ret = loadCore(ourXllPath.c_str());

    SetDllDirectory(NULL);

    theXllIsOpen = true;

    // We don't bother to hook xlEventCalculationEnded as this XLL event is not triggered
    // by programmatic recalc, but the COM event is hence is much more useful.
    // TODO: could this be called by xlOil.dll? I seem to remember not.
    if (ret == 1)
    {
      xloil::tryCallExcel(msxll::xlEventRegister,
        "xloHandleCalculationCancelled", msxll::xleventCalculationCanceled);
    }
  }
  catch (const std::exception& e )
  {
    logError(e.what());
  }
  return 1; // We alway return 1, even on failure.
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  try
  {
    if (theXllIsOpen)
      xloil::coreAutoClose(ourXllPath.c_str());
    theXllIsOpen = true;
  }
  catch (...)
  {
  }
  return 1;
}

// Temporarily removed as it's not adding value at the moment
/*
XLO_ENTRY_POINT(msxll::xloper12*) xlAddInManagerInfo12(msxll::xloper12* xAction)
{
  // This function can be called without the add-in loaded, so we avoid using
  // any xlOil functionality

  int action = 0;
  switch (xAction->xltype)
  {
  case msxll::xltypeNum:
    action = (int)xAction->val.num;
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
*/

XLO_ENTRY_POINT(void) xlAutoFree12(msxll::xloper12* pxFree)
{
  try
  {
    delete (xloil::ExcelObj*)pxFree;
  }
  catch (...)
  {
  }
}

XLO_ENTRY_POINT(int) xloHandleCalculationCancelled()
{
  if (theXllIsOpen)
    onCalculationCancelled();
  return 1;
}