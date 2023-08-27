#include <xloil/State.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/LogWindow.h>
#include <xloil/StaticRegister.h>
#include <xloil/StringUtils.h>
#include <xloil/ExcelThread.h>
#include <functional>
#include <filesystem>

namespace xloil { class RegisteredWorksheetFunc; }

namespace XllInfo
{
  inline static HINSTANCE dllHandle = NULL;
  inline static std::wstring xllName;
  inline static std::wstring xllPath;
}

// These functions are defined in XLO_DECLARE_ADDIN
void autoOpenDefinedInMacro();
void autoCloseDefinedInMacro();
std::wstring addInManagerInfo();

inline static bool theXllIsOpen = false;

void setDllPath(HMODULE handle)
{
  try
  {
    XllInfo::xllPath = xloil::captureWStringBuffer(
      [handle](auto* buf, auto len)
      {
        return GetModuleFileName(handle, buf, (DWORD)len);
      });

    if (!XllInfo::xllPath.empty())
    {
      XllInfo::xllName = std::filesystem::path(XllInfo::xllPath).filename();
      return;
    }
  }
  catch (...)
  {
  }

  OutputDebugStringW(L"xlOil_Loader: Could not determine XLL location");
}


namespace xloil
{
  template<class T>
  struct RegisterAddin
  {
    std::unique_ptr<T> theAddin;
    std::vector<std::shared_ptr<const RegisteredWorksheetFunc>> theFunctions;

    ~RegisterAddin()
    {
      if (theXllIsOpen)
      {
        XLO_ERROR(L"RegisterAddin: teardown initiated before xlAutoClose for {}", XllInfo::xllPath);
        autoClose();
      }
    }

    void autoOpen()
    {
      try
      {
        Environment::initAppContext();
          
        // A statically linked addin is its own core
        Environment::setCoreHandle(XllInfo::dllHandle);

        // Handle this event even if we have no registered async functions as they could be 
        // dynamically registered later
        tryCallExcel(msxll::xlEventRegister,
          "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);

        // Do this safely in single-thread mode
        initMessageQueue(Environment::excelProcess().hInstance);

        lauch();

        Environment::registerIntellisense(XllInfo::xllPath.c_str());

        std::wstring errorMessages;
        theFunctions = xloil::detail::registerStaticFuncs(XllInfo::xllName.c_str(), errorMessages);
        if (!errorMessages.empty())
          loadFailureLogWindow(XllInfo::dllHandle, errorMessages.c_str(), true);

        theXllIsOpen = true;
      }
      catch (const std::exception& e)
      {
        loadFailureLogWindow(XllInfo::dllHandle, utf8ToUtf16(e.what()), true);
      }
    }

    void lauch()
    {
      theAddin.reset(new T());
    }

    void autoClose()
    {
      if (theXllIsOpen)
        theAddin.reset();

      theFunctions.clear();
      theXllIsOpen = false;
    }
  };
}

XLO_ENTRY_POINT(int) DllMain(
  _In_ HINSTANCE hinstDLL,
  _In_ DWORD     fdwReason,
  _In_ LPVOID    /*lpvReserved*/
)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
  {
    XllInfo::dllHandle = hinstDLL;
    setDllPath(hinstDLL);
  }
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
    autoOpenDefinedInMacro();
  }
  catch (...) 
  {}
  return 1; // We alway return 1, even on failure.
}

XLO_ENTRY_POINT(int) xlAutoClose(void)
{
  try
  {
    autoCloseDefinedInMacro();
  }
  catch (...)
  {}
  return 1;
}

XLO_DEFINE_FREE_CALLBACK()

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

XLO_ENTRY_POINT(msxll::XLOPER12*) xlAddInManagerInfo12(msxll::XLOPER12* xAction)
{
  using namespace msxll;

  static XLOPER12 xIntAction, xInfo = { 1.0, xltypeNum }, xType = { 1.0, xltypeInt };
  xType.val.w = xltypeInt;

  Excel12(msxll::xlCoerce, &xIntAction, 2, xAction, xType);

  if (xInfo.xltype == xltypeStr)
    delete[] xInfo.val.str.data;

  try
  {
    auto info = addInManagerInfo();
    if (xIntAction.val.w == 1 && !info.empty())
    {
      xInfo.xltype = xltypeStr;
      xInfo.val.str.data = xloil::PString(info).release();
    }
    return &xInfo;
  }
  catch (...)
  {}

  xInfo.xltype = xltypeErr;
  xInfo.val.err = xlerrValue;

  return &xInfo;
}

XLO_ENTRY_POINT(int) xlAutoAdd()
{
  try
  {
    if (theXllIsOpen)
      xloil::Event::XllAdd().fire(XllInfo::xllName.c_str());
  }
  catch (...)
  { }
  return 1;
}

XLO_ENTRY_POINT(int) xlAutoRemove()
{
  try
  {
    if (theXllIsOpen)
      xloil::Event::XllRemove().fire(XllInfo::xllName.c_str());
  }
  catch (...)
  { }
  return 1;
}

namespace detail
{
  // Some ugly SFINAE to check if the an addInManagerInfo method is present
  auto callAddInManagerInfo(void*)
  {
    return std::wstring();
  }

  template<class T, std::enable_if_t<std::is_constructible_v<decltype(T::addInManagerInfo())>, bool> = true>
  auto callAddInManagerInfo(T*)
  {
    return T::addInManagerInfo();
  }
}

// Break this macro into two as the requirements for the Core addin are slightly
// different to user-built XLLs

#define _XLO_DECLARE_ADDIN_IMPL(T) \
  namespace { T theRegistedAddin; } \
  void autoOpenDefinedInMacro()  { theRegistedAddin.autoOpen(); } \
  void autoCloseDefinedInMacro() { theRegistedAddin.autoClose(); } \
  std::wstring addInManagerInfo() { return ::detail::callAddInManagerInfo((T*)nullptr); }

#define XLO_DECLARE_ADDIN(T) _XLO_DECLARE_ADDIN_IMPL(RegisterAddin<T>)