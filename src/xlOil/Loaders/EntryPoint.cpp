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
}

namespace xloil
{
  XLOIL_EXPORT int coreAutoOpen(const wchar_t* xllPath) noexcept
  {
    try
    {
      InXllContext xllContext;
      
      State::initAppContext(theCoreModuleHandle);
   
      bool firstLoad = openXll(xllPath);

      excelApp(); // Creates the COM connection
      initMessageQueue();

      return firstLoad ? 1 : 0;
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
