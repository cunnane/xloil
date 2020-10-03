#include <xloil/Ribbon.h>
#include <xloil/Log.h>
#include <xloil/ApiCall.h>

using namespace xloil;

  namespace
  {
    void ribbonHandler(const RibbonControl& ctrl)
    {
      XLO_TRACE(L"Ribbon action on {0}, {1}", ctrl.Id, ctrl.Tag);
    };

    std::shared_ptr<IComAddin> theComAddin;
  }

  void xllOpen(void* hInstance)
  {
    xllOpenComCall([]()
    {
      theComAddin = makeComAddin(L"TestXlOil");

      IComAddin::Handlers handlers;
      handlers[L"conBoldSub"] = ribbonHandler;
      handlers[L"conItalicSub"] = ribbonHandler;
      handlers[L"conUnderlineSub"] = ribbonHandler;

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
					      <button id="customButton3" label="ConUnderline" size="large" onAction="conUnderlineSub" imageMso="Underline" />
				      </group>
			      </tab>
		      </tabs>
	      </ribbon>
      </customUI>
      )", handlers);

      theComAddin->connect();

      theComAddin->ribbonInvalidate();
      theComAddin->ribbonActivate(L"customTab");
    });
  }
  void xllClose()
  {
    theComAddin.reset();
  }
