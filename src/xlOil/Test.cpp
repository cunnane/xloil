#include "Events.h"
#include "internal/FuncRegistry.h"
#include "ExcelObj.h"
#include "ExcelCall.h"
#include "Interface.h"
#include "ComInterface/Connect.h"
#include <xloil/StaticRegister.h>
#include <future>
using std::wstring;
using namespace msxll;

#ifdef _DEBUG

namespace xloil
{
  ExcelObj* testCallback(FuncInfo* info, const ExcelObj**)
  {
    return (new ExcelObj(wstring(info->name) + L" says hi"))->toExcel();
  }
}

extern "C" XLOIL_EXPORT XLOIL_XLOPER* WINAPI oilFoo(XLOIL_XLOPER* arg)
{
  XLOIL_XLOPER* args[1];
  args[0] = arg;
  return (new xloil::ExcelObj("oilFoo says hi"))->toExcel();
}

extern "C" __declspec(dllexport) XLOIL_XLOPER* WINAPI CallerExample(void)
{
  XLOPER12 xRes, xSheetName;

  Excel12(xlfCaller, &xRes, 0);
  Excel12(xlSheetNm, &xSheetName, 1, (LPXLOPER12)&xRes);

  return new xloil::ExcelObj(1);
}

namespace xloil
{

  struct DoRegister
  {
    static void run()
    {
      {
        auto info = std::make_shared<FuncInfo>(); 
        info->name = L"SomeArgs";
        info->args.push_back(FuncArg(L"Foo", L"Help"));
        registerFunc(info, &testCallback, info);
      }
      {
        auto info = std::make_shared<FuncInfo>();
        info->name = L"NoArgs";
        registerFunc(info, &testCallback, info);
      }
      {
        auto info2 = std::make_shared<FuncInfo>();
        info2->name = L"Foo";
        info2->args.push_back(FuncArg(L"Foo", L"Help"));
        registerFunc(info2, &testCallback, info2);
      }
      {
        auto info2 = std::make_shared<FuncInfo>(); 
        info2->name = L"CallerExample";
        registerFunc(info2, "CallerExample", Core::theCoreName());
      }
    }
  };


  struct RegisterMe
  {
    RegisterMe()
    {
      static auto handler = xloil::Event_AutoOpen() += []() { DoRegister::run(); };
    }
  } theInstance;

  XLO_ENTRY_POINT(void) xloAsyncTest(ExcelObj* asyncHandle, ExcelObj* arg)
  {
    try
    {
      callExcel(xlAsyncReturn, *asyncHandle, "Hello");
    }
    catch (...)
    {
    }
  }
  XLO_REGISTER(xloAsyncTest)
    .help(L"nope")
    .arg(L"foo", L"does foo")
    .async();
}

#endif

