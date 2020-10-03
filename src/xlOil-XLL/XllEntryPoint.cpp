#include <xlOil/XlCallSlim.h>
#include <xlOil/ExportMacro.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/StaticRegister.h>
#include <xlOil/Events.h>
#include <xlOil/State.h>
#include <filesystem>

namespace
{
  static HMODULE theModuleHandle = nullptr;
  std::vector<std::shared_ptr<const xloil::RegisteredFunc>> theFunctions;
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
  _In_ LPVOID    lpvReserved
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
    xloil::State::initAppContext(theModuleHandle);

    xllOpen(theModuleHandle);

    auto xllName = std::filesystem::path(
      xloil::callExcel(msxll::xlGetName).toString()).filename();
    theFunctions = xloil::registerStaticFuncs(xllName.c_str());

    // If you don't use native async functions, you don't need this
    xloil::tryCallExcel(msxll::xlEventRegister,
      "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);

    theXllIsOpen = true;
  }
  catch (...)
  {}
  return 1; // We alway return 1, even on failure.
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  try
  {
    if (theXllIsOpen)
      xllClose();
    
    theFunctions.clear();
    theXllIsOpen = false;
  }
  catch (...)
  {}
  return 1;
}

XLO_ENTRY_POINT(void) xlAutoFree12(msxll::xloper12* pxFree)
{
  try
  {
    delete (xloil::ExcelObj*)pxFree;
  }
  catch (...)
  {}
}

XLO_ENTRY_POINT(int) xlHandleCalculationCancelled()
{
  try
  {
    if (theXllIsOpen)
      xloil::Event::CalcCancelled().fire();
  }
  catch (...)
  {}
  return 1;
}