#include <xloil/State.h>
#include <xloil/ExcelCall.h>
#include <xloil/Events.h>
#include <xloil/LogWindow.h>
#include <xloil/StaticRegister.h>
#include <functional>
#include <filesystem>

namespace xloil
{
  class RegisteredFunc;

  namespace XllInfo
  {
    inline static HMODULE moduleHandle = NULL;
    inline static std::wstring xllName;
    inline static std::wstring xllPath;
  }
  namespace detail
  {
    void dllLoad();
    void autoOpen();
    void autoClose();
    std::wstring addInManagerInfo();

    inline static bool theXllIsOpen = false;

    void setDllPath(HMODULE handle)
    {
      try
      {
        wchar_t path[4 * MAX_PATH]; // TODO: may not be long enough!!!
        auto size = GetModuleFileName(handle, path, sizeof(path));
        if (size > 0)
        {
          XllInfo::xllPath = std::wstring(path, path + size);
          XllInfo::xllName = std::filesystem::path(XllInfo::xllPath).filename();
          return;
        }
      }
      catch (...)
      {}
      OutputDebugStringW(L"xloil_Loader: Could not determine XLL location");
    }

    template<class T>
    struct Reg
    {
      inline static std::unique_ptr<T> theAddin;
      inline static std::vector<std::shared_ptr<const RegisteredFunc>> theFunctions;

      static void autoOpen()
      {
        try
        {
          State::initAppContext();

          // TODO: check if we have registered async functions
          tryCallExcel(msxll::xlEventRegister,
            "xlHandleCalculationCancelled", msxll::xleventCalculationCanceled);

          theAddin.reset(new T());

          std::wstring errorMessages;
          theFunctions = xloil::registerStaticFuncs(XllInfo::xllName.c_str(), errorMessages);
          if (!errorMessages.empty())
            writeLogWindow(errorMessages.c_str());

          // Do now this safely in single-thread mode
          initMessageQueue();

          theXllIsOpen = true;
        }
        catch (const std::exception& e)
        {
          writeLogWindow(e.what());
        }
      }
      static void autoClose()
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
      XllInfo::moduleHandle = hinstDLL;
      detail::setDllPath(hinstDLL);
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
      detail::autoOpen();
    }
    catch (...) 
    {}
    return 1; // We alway return 1, even on failure.
  }

  XLO_ENTRY_POINT(int) xlAutoClose(void)
  {
    try
    {
      detail::autoClose();
    }
    catch (...)
    {}
    return 1;
  }

  XLO_ENTRY_POINT(void) xlAutoFree12(ExcelObj* pxFree)
  {
    try
    {
      delete pxFree;
    }
    catch (...)
    { }
  }

  XLO_ENTRY_POINT(int) xlHandleCalculationCancelled()
  {
    try
    {
      if (detail::theXllIsOpen)
        Event::CalcCancelled().fire();
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
      delete[] xInfo.val.str;

    try
    {
      auto info = detail::addInManagerInfo();
      if (xIntAction.val.w == 1 && !info.empty())
      {
        xInfo.xltype = xltypeStr;
        xInfo.val.str = PString<>(info).release();
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
      if (detail::theXllIsOpen)
        Event::XllAdd().fire(XllInfo::xllName.c_str());
    }
    catch (...)
    { }
    return 1;
  }

  XLO_ENTRY_POINT(int) xlAutoRemove()
  {
    try
    {
      if (detail::theXllIsOpen)
        Event::XllRemove().fire(XllInfo::xllName.c_str());
    }
    catch (...)
    { }
    return 1;
  }

  namespace detail
  {
    // This is all really really horrible, the template SFINAE is so
    // painfully awkward.

    template<class T>
    auto callAutoOpen(T*, void*) 
    {
      Reg<T>::autoOpen();
    }
    template<class T, std::enable_if_t<std::is_void<decltype(T::autoOpen())>::value, bool> = true>
    auto callAutoOpen(T*, T*)
    {
      T::autoOpen();
    }

    template<class T>
    auto callAutoClose(T*, void*)
    {
      Reg<T>::autoClose();
    }
    template<class T, std::enable_if_t<std::is_void<decltype(T::autoClose())>::value, bool> = true>
    auto callAutoClose(T*, T*)
    {
      T::autoClose();
    }

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

#define XLO_DECLARE_ADDIN(T) \
  void detail::autoOpen()  { detail::callAutoOpen((T*)nullptr, (T*)nullptr); } \
  void detail::autoClose() { detail::callAutoClose((T*)nullptr, (T*)nullptr); } \
  std::wstring detail::addInManagerInfo() { return detail::callAddInManagerInfo((T*)nullptr); }
}