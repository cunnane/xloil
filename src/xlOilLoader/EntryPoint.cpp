#include <xloilHelpers/WindowsSlim.h>
#include <xloilHelpers/Environment.h>
#include <xloilHelpers/Settings.h>
#include <xlOil/XlCallSlim.h>
#include "xloil/EntryPoint.h"
#include "xloil/ExportMacro.h"
#include "xloil/Log.h"
#include "xloil/ExcelCall.h"
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
  static wstring ourXllDir;
  static bool theXllHasClosed = false;

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
      ourXllDir = fs::path(ourXllPath).remove_filename().wstring();
      return true;
    }
    catch (...)
    {
      return false;
    }
  }

  std::fstream theLogFile;

  /// <summary>
  /// Very cheap log file to catch startup errors before
  /// the core dll can initialise spdlog.
  /// </summary>
  void logError(const std::string& err)
  {
    OutputDebugStringA(err.c_str());
    if (!theLogFile.good())
      theLogFile = std::fstream(fs::path(ourXllPath).replace_extension("log"));
    theLogFile << err << std::endl;
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
  switch (dliNotify)
  {
  case dliFailGetProc:
    throw std::runtime_error("Unable to find procedure: "s
      + pdli->dlp.szProcName + " in " + pdli->szDll);
  default:
    throw std::runtime_error("Unable to load library: "s + pdli->szDll);
  }
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

/// <summary>
/// Looks for the core xlOil.dll
/// </summary>
fs::path wheresWally(const wchar_t* dll)
{
  // Check already loaded
  if (GetModuleHandle(dll) != 0)
    return fs::path();

  // Check same directory as XLL
  fs::path path = fs::path(ourXllPath).remove_filename();
  if (fs::exists(path / dll))
    return path;

  // Check if the addin is installed
  auto excelVersion = getExcelVersion();
  wstring addinPath;
  size_t iAddin = 0; // Small bug, first OPEN key does not have a number
  while (getWindowsRegistryValue(
    L"HKCU",
    formatWStr(L"Software\\Microsoft\\Office\\%d.0\\Excel\\Options\\OPEN%s",
      excelVersion, iAddin > 0 ? std::to_wstring(iAddin) : wstring()),
    addinPath))
  {
    for (auto& c : addinPath) c = toupper(c);
    if (addinPath.find(dll) != wstring::npos)
      return fs::path(addinPath).remove_filename();
  }

  return fs::path();
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
    // We need to find xloil.dll. The search order is
    // 0) Check already loaded
    // 1) Look in same dir as XLL
    // 2) Look in Excel addins in registry for xlOil.xll
    // 3) Look in %APPDATA%/xloil settings file for Environment

    vector<shared_ptr<PushEnvVar>> environmentVariables;

    auto settings = findSettingsFile(L"xlOil.dll");

    // Check if the settings file contains an Environment block
    if (settings)
    {
      auto environment = Settings::environmentVariables(
        (*settings)["Addin"]);
      
      for (auto&[key, val] : environment)
      {
        auto value = expandWindowsRegistryStrings(
          expandEnvironmentStrings(val));

        environmentVariables.emplace_back(
          std::make_shared<PushEnvVar>(key.c_str(), value.c_str()));
      }
    }

    auto dllPath = wheresWally(xloil_dll);
    if (!dllPath.empty())
      SetDllDirectory(dllPath.c_str());
    int ret = xloil::coreAutoOpen(ourXllPath.c_str());
    SetDllDirectory(NULL);

    if (ret > 0)
    {
      // xleventCalculationEnded not hooked as the XLL event is not triggered
      // by programmatic recalc, but the COM event (more usefully) is
      xloil::tryCallExcel(msxll::xlEventRegister,
        "xloHandleCalculationCancelled", msxll::xleventCalculationCanceled);
    }
  }
  catch (const std::exception& e)
  {
    logError(e.what());
  }
  catch (...)
  {
  }
  return 1;
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  try
  {
    xloil::coreAutoClose(ourXllPath.c_str());
    theXllHasClosed = true;
  }
  catch (...)
  {
  }
  return 1;
}

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
  // An excel "undocumented feature" is to call XLL event handlers
  // after calling xlAutoClose, which is clearly not ideal.
  // This seems to happen when Excel is closing and asks the user 
  // to save the workbook.
  if (!theXllHasClosed)
    onCalculationCancelled();
  return 1;
}