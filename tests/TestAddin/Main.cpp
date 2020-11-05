#include <xloil/Ribbon.h>
#include <xloil/Log.h>
#include <xloil/ApiCall.h>
#include <xloil/DynamicRegister.h>
#include <xloil/Async.h>

#include <map>
using namespace xloil;
using std::wstring;
using std::shared_ptr;

namespace
{
  void ribbonHandler(const RibbonControl& ctrl, VARIANT* ret, int nArgs, tagVARIANT** args)
  {
    XLO_TRACE(L"Ribbon action on {0}, {1}", ctrl.Id, ctrl.Tag);
  };

  shared_ptr<IComAddin> theComAddin;
  std::list<shared_ptr<RegisteredFunc>> theFuncs;
}

  void xllOpen(void* hInstance)
  {
    theFuncs.push_back(RegisterLambda<>(
      [](const ExcelObj& arg1, const ExcelObj& arg2)
      {
          return returnValue(7);
      })
      .name(L"testDynamic")
      .arg(L"Arg1")
      .registerFunc());
    theFuncs.push_back(RegisterLambda<void>(
      [](const FuncInfo& info, const ExcelObj& arg1, const AsyncHandle& handle)
      {
        handle.returnValue(8);
      })
      .name(L"testDynamicAsync")
      .arg(L"Arg1")
      .registerFunc());

    xllOpenComCall([]()
    {
      theComAddin = makeComAddin(L"TestXlOil");

      std::map<wstring, IComAddin::RibbonCallback> handlers;
      handlers[L"conBoldSub"] = ribbonHandler;
      handlers[L"conItalicSub"] = ribbonHandler;
      handlers[L"comboChange"] = ribbonHandler;
      auto mapper = [=](const wchar_t* name) mutable { return handlers[name]; };

      theComAddin->setRibbon(LR"(
      <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
	      <ribbon>
		      <tabs>
			      <tab id="customTab" label="xlOilTest" insertAfterMso="TabHome">
				      <group idMso="GroupClipboard" />
				      <group idMso="GroupFont" />
				      <group id="customGroup" label="MyButtons">
					      <button id="customButton1" label="ConBold" size="large" onAction="conBoldSub" imageMso="Bold" />
					      <button id="customButton2" label="ConItalic" size="large" onAction="conItalicSub" imageMso="Italic" />
					      <comboBox id="comboBox" label="Combo Box" onChange="comboChange">
                 <item id="item1" label="Item 1" />
                 <item id="item2" label="Item 2" />
                 <item id="item3" label="Item 3" />
               </comboBox>
				      </group>
			      </tab>
		      </tabs>
	      </ribbon>
      </customUI>
      )", mapper);

      theComAddin->connect();

      theComAddin->ribbonInvalidate();
      theComAddin->ribbonActivate(L"customTab");
    });
  }

  void xllClose()
  {
    theComAddin.reset();
    theFuncs.clear();
  }
