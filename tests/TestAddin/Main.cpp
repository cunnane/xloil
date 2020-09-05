#include <xloil/Interface.h>
#include <xloil/Log.h>

namespace xloil
{
  namespace Test
  {
    void ribbonHandler(const RibbonControl& ctrl)
    {
      XLO_TRACE(L"Ribbon action on {0}, {1}", ctrl.Id, ctrl.Tag);
    };

    std::shared_ptr<IComAddin> theComAddin;

    XLO_PLUGIN_INIT(AddinContext* addin, const PluginContext& plugin)
    {
      linkLogger(addin, plugin);

      if (plugin.action == PluginContext::Load)
      {
        theComAddin = makeComAddin(L"TestXlOil");

        std::map<std::wstring, std::function<void(const RibbonControl&)>> handlers;
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
      } 
      else if (plugin.action == PluginContext::Unload)
      {
        theComAddin.reset();
      }
      return 0;
    }
  }
}

