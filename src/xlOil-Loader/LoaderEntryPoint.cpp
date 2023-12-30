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
#define TOML_ABI_NAMESPACES 0
#include <toml++/toml.h>
#include <filesystem>
#define DELAYIMP_INSECURE_WRITABLE_HOOKS
#include <delayimp.h>
#include <fstream>
#include <sstream>
#include <string>

namespace fs = std::filesystem;

using std::string;
using std::wstring;
using namespace xloil;
using std::vector;
using std::shared_ptr;
using namespace std::string_literals;

namespace
{
  void writeToStartUpLog(const string& msg, bool openWindow=false) noexcept
  {
    OutputDebugStringA(msg.c_str());
    loadFailureLogWindow(XllInfo::dllHandle, utf8ToUtf16(msg), openWindow);
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
  writeToStartUpLog(msg);
  return nullptr;
}

extern "C" PfnDliHook __pfnDliFailureHook2 = nullptr;
namespace
{
  // Name of core dll. Not sure if there's an automatic way to get this
  // maybe some build env vars?
  constexpr wchar_t* const xloil_dll = L"xlOil.dll";

  void loadEnvironmentBlock(const toml::table& settings)
  {
    auto environment = Settings::environmentVariables(settings[XLOIL_SETTINGS_ADDIN_SECTION]);

    for (auto&[key, val] : environment)
    {
      auto value = expandWindowsRegistryStrings(
        expandEnvironmentStrings(val));

      SetEnvironmentVariable(key.c_str(), value.c_str());
    }
  }

  auto findAllCoreDllImports()
  {
    // If the delay load fails, it will throw a SEH exception, so we must use
    // __try/__except to avoid this crashing Excel.
    auto previousHook = __pfnDliFailureHook2;
    __pfnDliFailureHook2 = &delayLoadFailureHook;
    __try
    {
      __HrLoadAllImportsForDll("xlOil.dll"); // This is CASE SENSITIVE!
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

struct xlOilCoreAddin
{
  std::vector<std::shared_ptr<const RegisteredWorksheetFunc>> theFunctions;
  std::vector<std::string> _loadingMessages;

  template<class...Args>
  void log(const char* fmt, Args...args)
  {
    _loadingMessages.emplace_back(formatStr(fmt, std::forward<Args>(args)...));
  }

  void autoOpen()
  {
    try
    {
      using XllInfo::xllPath;
      
      std::unique_ptr<PushDllDirectory> setDllDir;

      // First we try to load a settings file to see if it tells us
      // to run a startup trace
      const auto settings = findSettingsFile(xllPath.c_str());
      
      std::error_code fsErr;
      if (settings)
        log("Found ini file at '%s'", settings->source().path->c_str());

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
        _loadingMessages.emplace_back("xlOil.dll already loaded");
      }
      else if (fs::exists(ourXllDir / xloil_dll, fsErr)) // Check same directory as XLL
      {
        log("Found xlOil.dll in xll directory. Calling SetDllDirectory('%s')",
            ourXllDir.string().c_str());
        setDllDir.reset(new PushDllDirectory(ourXllDir.c_str()));
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
            log("Found xloil.ini at '%s'", coreSettings->source().path->c_str());
            loadEnvironmentBlock(*coreSettings);
          }
        }
      }

      log("Environment PATH=%s", getEnvironmentVar("PATH").c_str());
    
      if (!findAllCoreDllImports())
      {
        SetDllDirectory(NULL);
        throw std::runtime_error("Failed to load xlOil.dll, check XLOIL_PATH in ini file");
      }


      Environment::initAppContext();

      // Now we can use the normal logger!
      XLO_DEBUG("xlOil Core: Opening");
      auto ret = xloil::coreAutoOpenHandler(XllInfo::xllPath.c_str());

      if (ret == 1)
      {
        tryCallExcel(msxll::xlEventRegister,
          "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);
      }

      theXllIsOpen = true;
    }
    catch (const std::exception& e)
    {
      for (auto& msg : _loadingMessages)
        writeToStartUpLog(msg);

      writeToStartUpLog(e.what(), true);
    }

    _loadingMessages.clear();
  }

  void autoClose()
  {
    XLO_DEBUG("xlOil Core: Closing");
    xloil::coreAutoCloseHandler(XllInfo::xllPath.c_str());

    theFunctions.clear();
    theXllIsOpen = false;
  }

  auto addInManagerInfo()
  {
    return std::wstring(L"xlOil Core");
  }
};

_XLO_DECLARE_ADDIN_IMPL(xlOilCoreAddin);