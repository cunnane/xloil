#include <xloil/State.h>
#include <xlOil/WindowsSlim.h>
#include <xloil/Throw.h>
#include <xloil/ExcelCall.h>
#include <xloil/AppObjects.h>
#include "Intellisense.h"
#include <filesystem>

namespace xloil
{
  namespace
  {
    static HMODULE theCoreModuleHandle = nullptr;

    static std::wstring ourDllName;
    static std::wstring ourDllPath;
    static Environment::ExcelProcessInfo ourExcelState;

    // TODO: make this startup stuff noexcept?
    void setDllPath(HMODULE handle)
    {
      ourDllPath = captureWStringBuffer(
        [handle](auto* buf, auto len)
        {
          return GetModuleFileName(handle, buf, (DWORD)len);
        });
      ourDllName = std::filesystem::path(ourDllPath).filename();
    }

    int getExcelVersion()
    {
      // https://github.com/MicrosoftDocs/office-developer-client-docs/blob/...
      // master/docs/excel/calling-into-excel-from-the-dll-or-xll.md
      auto versionStr = callExcel(msxll::xlfGetWorkspace, 2);
      return std::stoi(versionStr.toString());
    }

    HINSTANCE getExcelHInstance()
    {
      auto instPtr = callExcel(msxll::xlGetInstPtr);
      return (HINSTANCE)instPtr.val.bigdata.h.hdata;
    }

    HWND getExcelHWnd()
    {
      auto hwnd = callExcel(msxll::xlGetHwnd);
      return (HWND)IntToPtr(hwnd.val.w);
    }
  }

  namespace Environment
  {
    XLOIL_EXPORT void initAppContext()
    {
      ourExcelState.version = getExcelVersion();
      ourExcelState.hInstance = getExcelHInstance();
      ourExcelState.hWnd = (long long)getExcelHWnd();
      ourExcelState.mainThreadId = GetCurrentThreadId();
    }
    void setCoreHandle(void* coreHInstance)
    {
      if (theCoreModuleHandle)
        return;

      theCoreModuleHandle = (HMODULE)coreHInstance;
      setDllPath(theCoreModuleHandle);
    }
    void* coreModuleHandle() noexcept
    {
      return theCoreModuleHandle;
    }
   
    const wchar_t* coreDllPath() noexcept
    {
      return ourDllPath.c_str();
    }
    const wchar_t* coreDllName() noexcept
    {
      return ourDllName.c_str();
    }

    ExcelProcessInfo::ExcelProcessInfo()
      : version(0)
      , hInstance(nullptr)
      , hWnd(0)
      , mainThreadId(0)
    {}

    XLOIL_EXPORT const ExcelProcessInfo& excelProcess() noexcept
    {
      return ourExcelState;
    }

    void registerIntellisense(const wchar_t* xllPath)
    {
      registerIntellisenseHook(xllPath);
    }
  }
}