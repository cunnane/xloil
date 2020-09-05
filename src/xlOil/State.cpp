#include <xloil/State.h>
#include <xlOil/WindowsSlim.h>
#include <xloil/Throw.h>
#include <xloil/ExcelCall.h>
#include <COMInterface/Connect.h>

namespace xloil
{
  namespace
  {
    static HMODULE theCoreModuleHandle = nullptr;

    static const wchar_t* ourDllName = nullptr;
    static wchar_t ourDllPath[4 * MAX_PATH]; // TODO: may not be long enough!!!
    static int ourExcelVersion = 0;
    static HINSTANCE ourExcelHInstance = nullptr;
    static DWORD ourMainThread = 0;

    void setDllPath(HMODULE handle)
    {
      auto size = GetModuleFileName(handle, ourDllPath, sizeof(ourDllPath));
      if (size == 0)
        XLO_THROW(L"Could not determine XLL location: {}", xloil::writeWindowsError());
      ourDllName = wcsrchr(ourDllPath, L'\\') + 1;
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
  }

  namespace State
  {
    void initAppContext(void* coreHInstance)
    {
      theCoreModuleHandle = (HMODULE)coreHInstance;
      setDllPath(theCoreModuleHandle);
      ourMainThread = GetCurrentThreadId();
      ourExcelVersion = getExcelVersion();
      ourExcelHInstance = getExcelHInstance();
    }

    void* coreModuleHandle() noexcept
    {
      return theCoreModuleHandle;
    }
   
    const wchar_t* corePath() noexcept
    {
      return ourDllPath;
    }
    const wchar_t* coreName() noexcept
    {
      return ourDllName;
    }
    int excelVersion() noexcept
    {
      return ourExcelVersion;
    }
    size_t mainThreadId() noexcept
    {
      return ourMainThread;
    }
    void* excelHInstance() noexcept
    {
      return ourExcelHInstance;
    }
    Excel::_Application& excelApp() noexcept
    {
      return COM::excelApp();
    }
  }
}