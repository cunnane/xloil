#include <xloilHelpers/Environment.h>
#include <xloilHelpers/Settings.h>
#include <xlOil/XlCallSlim.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Events.h>
#include <xlOil-XLL/LogWindow.h>
#include <tomlplusplus/toml.hpp>
#include <filesystem>
#define DELAYIMP_INSECURE_WRITABLE_HOOKS
#include <delayimp.h>
#include <fstream>
#include <sstream>

namespace fs = std::filesystem;

using std::string;
using std::wstring;
using namespace xloil;
using std::vector;
using std::shared_ptr;
using namespace std::string_literals;
using namespace xloil::Helpers;

namespace
{
  static HMODULE theModuleHandle = nullptr;
  static wstring ourXllPath;
  static wstring ourLogFilePath;

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
      ourLogFilePath = fs::path(ourXllPath).replace_extension("log");
      return true;
    }
    catch (...)
    {
      return false;
    }
  }

  void writeLog(const std::string& msg)
  {
    writeLogWindow(msg.c_str());
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
  writeLog(msg);
  return nullptr;
}

extern "C" PfnDliHook __pfnDliFailureHook2 = nullptr;

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

int loadCore(const wchar_t* xllPath)
{
  // If the delay load fails, it will throw a SEH exception, so we must use
  // __try/__except to avoid this crashing Excel.
  int ret = -1;
  auto previousHook = __pfnDliFailureHook2;
  __pfnDliFailureHook2 = &delayLoadFailureHook;
  __try
  {
    ret = xloil::autoOpenHandler(xllPath);
  }
  __except (EXCEPTION_EXECUTE_HANDLER)
  {
    // TODO: add GetExceptionCode(), without using string
    writeLogWindow("Failed to load xlOil.dll, check XLOIL_PATH in ini file");
  }
  __pfnDliFailureHook2 = previousHook;
  return ret;
}

void xllOpen(void* hInstance)
{
  try
  {
    setDllPath((HMODULE)hInstance);

    // We need to find xloil.dll. 
    const auto ourXllDir = fs::path(ourXllPath).remove_filename();

    const auto settings = findSettingsFile(ourXllPath.c_str());
    auto traceLoad = false;
    if (settings)
    {
      traceLoad = (*settings)["Addin"]["StartupTrace"].value_or(false);
      ourLogFilePath = Settings::logFilePath(*settings);
      if (traceLoad)
        writeLog(formatStr("Found ini file at '%s'", settings->source().path->c_str()));
    }
    if (GetModuleHandle(xloil_dll) != 0) // Is it already loaded?
    {
      if (traceLoad)
        writeLog("xlOil.dll already loaded!");
    }
    else if (fs::exists(ourXllDir / xloil_dll)) // Check same directory as XLL
    {
      if (traceLoad)
        writeLog(formatStr("Found xlOil.dll, using SetDllDirectory = '%s'", ourXllDir.string().c_str()));
      SetDllDirectory(ourXllDir.c_str());
    }
    else
    {
      // Load the environment block from our settings file (if it exists)
      if (settings)
        loadEnvironmentBlock(*settings);

      // If we aren't xloil.xll (where we would already have loaded xloil.ini)
      // look for xloil.ini and see if it contains an enviroment block
      if (_wcsicmp(fs::path(ourXllPath).filename().c_str(), L"xloil.xll") != 0)
      {
        auto coreSettings = findSettingsFile(
          fs::path(ourXllPath).replace_filename(xloil_dll).c_str());
        if (coreSettings)
        {
          writeLog(formatStr("Found xloil.ini at '%s'", coreSettings->source().path->c_str()));
          loadEnvironmentBlock(*settings);
        }
      }
    }

    if (traceLoad)
      writeLog(formatStr("Environment PATH=%s", getEnvVar("PATH").c_str()));

    auto ret = loadCore(ourXllPath.c_str());

    SetDllDirectory(NULL);
  }
  catch (const std::exception& e)
  {
    writeLog(e.what());
  }
}

void xllClose()
{
  xloil::autoCloseHandler(ourXllPath.c_str());
}


namespace
{
  // This bool is required due to apparent bugs in the XLL interface:
  // Excel may call XLL event handlers after calling xlAutoClose,
  // and it may call xlAutoClose without ever having called xlAutoOpen
  // This former to happen when Excel is closing and asks the user 
  // to save the workbook, the latter when removing an addin using COM
  // automation
  bool theXllIsOpen = false;
}

void xllOpen(void* hInstance);
void xllClose();

XLO_ENTRY_POINT(int) DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    /*lpvReserved*/
)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
    theModuleHandle = hinstDLL;

  return TRUE;
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
    xllOpen(theModuleHandle);
    
    xloil::tryCallExcel(msxll::xlEventRegister,
      "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);

    theXllIsOpen = true;
  }
  catch (...)
  {
  }
  return 1; // We alway return 1, even on failure.
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  try
  {
    if (theXllIsOpen)
      xllClose();

    theXllIsOpen = false;
  }
  catch (...)
  {
  }
  return 1;
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

XLO_ENTRY_POINT(int) xlHandleCalculationCancelled()
{
  try
  {
    if (theXllIsOpen)
      xloil::Event::CalcCancelled().fire();
  }
  catch (...)
  {
  }
  return 1;
}