#include <xlOil/Events.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/Interface.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/Log.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xlOilHelpers/WindowsSlim.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xlOil/State.h>
#include <xloil/ThreadControl.h>
#include <COMInterface/Connect.h>
#include <COMInterface/XllContextInvoke.h>


using std::wstring;
using std::string;
using std::vector;
using std::shared_ptr;

namespace
{
  static HMODULE theCoreModuleHandle = nullptr;
  static bool theCoreIsLoaded = false;
}

namespace xloil
{
  struct RetryAtStartup
  {
    void operator()()
    {
      try
      {
        connectCOM();
        excelApiCall([=]() { openXll(path.c_str()); }, QueueType::XLL_API);
      }
      catch (const ComConnectException& e)
      {
        Sleep(1000);
        excelApiCall(RetryAtStartup{ path }, QueueType::WINDOW | QueueType::ENQUEUE, 0);
      }
    }
    wstring path;
  };

  XLOIL_EXPORT int coreAutoOpen(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;

      State::initAppContext(theCoreModuleHandle);
      // A return val of 1 tells the XLL to hook XLL-api events. There may be
      // mulltiple XLLs, but we only want to hook the events once, when we load 
      // the core DLL.
      int retVal = 0;

      if (!theCoreIsLoaded)
      {
#if _DEBUG
        detail::loggerInitialise(spdlog::level::debug);
#else
        detail::loggerInitialise(spdlog::level::warn);
#endif
        openCore();

        theCoreIsLoaded = true;
        retVal = 1;
      }

      initMessageQueue();

      excelApiCall(RetryAtStartup{ wstring(xllPath) });

      return retVal;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Initialisation error: {0}", e.what());
    }
    return -1;
  }
  XLOIL_EXPORT int coreAutoClose(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;
      
      closeXll(xllPath);

      return 1;
    }
    catch (const std::exception& e)
    {
      XLO_ERROR("Finalisation error: {0}", e.what());
    }
    return 0;
  }

  XLOIL_EXPORT void onCalculationEnded() noexcept
  {
    try
    {
      InXllContext xllContext;
      xloil::Event::AfterCalculate().fire();
    }
    catch (...)
    {
    }
  }
  XLOIL_EXPORT void onCalculationCancelled() noexcept
  {
    try
    {
      InXllContext xllContext;
      xloil::Event::CalcCancelled().fire();
    }
    catch (...)
    {
    }
  }
}

XLO_ENTRY_POINT(int) DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    /*lpvReserved*/
)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
  {
    theCoreModuleHandle = hinstDLL;
  }
  return TRUE;
}

extern "C"  __declspec(dllexport) void* __stdcall XLOIL_STUB_NAME() 
{ 
  return nullptr; 
}
