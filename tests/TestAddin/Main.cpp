#include <xloil/ExcelUI.h>
#include <xloil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xloil/DynamicRegister.h>
#include <xloil/Async.h>
#include <xloil/XllEntryPoint.h>
#include <xloilHelpers/Environment.h>

#include <map>
using namespace xloil;
using std::wstring;
using std::shared_ptr;

namespace
{
  void ribbonHandler(const RibbonControl& ctrl, VARIANT* /*ret*/, 
                     int /*nArgs*/, tagVARIANT** /*args*/)
  {
    XLO_TRACE(L"Ribbon action on {0}, {1}", ctrl.Id, ctrl.Tag);
  };
}

struct MyAddin
{
  shared_ptr<IComAddin> theComAddin;
  std::list<shared_ptr<RegisteredWorksheetFunc>> theFuncs;

  MyAddin()
  {
    auto logger = loggerInitialise("warn");
    // It's not a great idea to put your log file in the same directory as the XLL
    // because if the XLL has been added to Excel's `XLSTART` folder, Excel will
    // attempt to open the log file when it is next started.
    loggerAddRotatingFileSink(logger,
      getEnvironmentVar(L"APPDATA") + L"\\xlOil\\" + XllInfo::xllName + L".log",
      "debug",
      1000);

    theFuncs.push_back(RegisterLambda<>(
      [](const ExcelObj& /*arg1*/, const ExcelObj& /*arg2*/)
      {
        return returnValue(7);
      })
      .name(L"testDynamic")
      .arg(L"Arg1")
      .registerFunc());
    theFuncs.push_back(RegisterLambda<void>(
      [](const FuncInfo& /*info*/, const ExcelObj& /*arg1*/, const AsyncHandle& handle)
      {
        handle.returnValue(8);
      })
      .name(L"testDynamicAsync")
      .arg(L"Arg1")
      .registerFunc());
    theFuncs.push_back(RegisterLambda<int>(
      [](const FuncInfo& /*info*/, const ExcelObj& /*arg1*/)
      {
        return 1;
      })
      .name(L"testDynamicAsync")
      .arg(L"Arg1")
      .command()
      .registerFunc());

    runComSetupOnXllOpen([this]()
    {
      theComAddin = makeComAddin(L"TestXlOil");

      std::map<wstring, IComAddin::RibbonCallback> handlers;
      handlers[L"conBoldSub"] = ribbonHandler;
      handlers[L"conItalicSub"] = ribbonHandler;
      handlers[L"comboChange"] = ribbonHandler;
      auto mapper = [=](const wchar_t* name) mutable { return handlers[name]; };

      theComAddin->connect(LR"(
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

      theComAddin->ribbonInvalidate();
      theComAddin->ribbonActivate(L"customTab");

      std::shared_ptr<ICustomTaskPane> taskPane(theComAddin->createTaskPane(L"xloil"));
      taskPane->setVisible(true);
    });
  }

  ~MyAddin()
  {
    theComAddin.reset();
    theFuncs.clear();
  }

  static wstring addInManagerInfo()
  {
    return wstring(L"xlOil Static Test");
  }
};

XLO_DECLARE_ADDIN(MyAddin);
