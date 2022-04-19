#include <xloilHelpers/Environment.h>
#include <xloilHelpers/Settings.h>
#include <xlOil/XlCallSlim.h>
#include <xlOil/Loaders/CoreEntryPoint.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/Events.h>
#include <xlOil/LogWindow.h>
#include <xloil/XllEntryPoint.h>
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

namespace
{
  void writeLog(const string& msg) noexcept
  {
    loadFailureLogWindow(XllInfo::dllHandle, utf8ToUtf16(msg));
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
    break;
  default:
    msg = formatStr("Unable to load library: %s", pdli->szDll);
  }
  writeLog(msg);
  return nullptr;
}

extern "C" PfnDliHook __pfnDliFailureHook2 = nullptr;
namespace
{
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

  auto loadCore()
  {
    // If the delay load fails, it will throw a SEH exception, so we must use
    // __try/__except to avoid this crashing Excel.
    auto previousHook = __pfnDliFailureHook2;
    __pfnDliFailureHook2 = &delayLoadFailureHook;
    __try
    {
      __HrLoadAllImportsForDll("xlOil.dll");
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
      // TODO: add GetExceptionCode(), without using string
      return false;
    }
    __pfnDliFailureHook2 = previousHook;
    return true;
  }
}

struct xlOilAddin
{
  static void autoOpen()
  {
    try
    {
      using XllInfo::xllPath;
      
      // First we try to load a settings file to see if it tells us
      // to run a startup trace
      const auto settings = findSettingsFile(xllPath.c_str());
      auto traceLoad = false;
      std::error_code fsErr;
      if (settings)
      {
        traceLoad = (*settings)["Addin"]["StartupTrace"].value_or(false);
        if (traceLoad)
          writeLog(formatStr("Found ini file at '%s'", settings->source().path->c_str()));
      }

      // Next we need to find xloil.dll. The strategy is
      //   1. Is it already loaded?
      //   2. Is it in the same directory as the XLL? Then use SetDllDirectory
      //   3. Apply any environment variables (in particular PATH) which
      //      are specifed in the ini file
      //   4. Look for xloil.ini and apply those env vars as well
      // Hope the above has setup the environment in the right way!
      const auto ourXllDir = fs::path(xllPath).remove_filename();
      if (GetModuleHandle(xloil_dll) != 0) // Is it already loaded?
      {
        if (traceLoad)
          writeLog("xlOil.dll already loaded!");
      }
      else if (fs::exists(ourXllDir / xloil_dll, fsErr)) // Check same directory as XLL
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
        if (_wcsicmp(fs::path(xllPath).filename().c_str(), L"xloil.xll") != 0)
        {
          auto coreSettings = findSettingsFile(
            fs::path(xllPath).replace_filename(xloil_dll).c_str());
          if (coreSettings)
          {
            if (traceLoad)
              writeLog(formatStr("Found xloil.ini at '%s'", coreSettings->source().path->c_str()));
            loadEnvironmentBlock(*settings);
          }
        }
      }

      if (traceLoad)
        writeLog(formatStr("Environment PATH=%s", getEnvVar("PATH").c_str()));
    
      if (!loadCore())
        writeLog("Failed to load xlOil.dll, check XLOIL_PATH in ini file");

      SetDllDirectory(NULL);

      State::initAppContext();

      detail::RegisterAddinBase<xlOilAddin>::theAddin.reset(new xlOilAddin());

      detail::theXllIsOpen = true;
    }
    catch (const std::exception& e)
    {
      writeLog(e.what());
    }
  }
  xlOilAddin() 
  {
    auto ret = xloil::autoOpenHandler(XllInfo::xllPath.c_str());

    if (ret == 1)
    {
      tryCallExcel(msxll::xlEventRegister,
        "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);
    }
  }
  ~xlOilAddin()
  {
    xloil::autoCloseHandler(XllInfo::xllPath.c_str());
  }
};

XLO_DECLARE_ADDIN(xlOilAddin);