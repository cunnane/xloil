#include <xloil/Plugin.h>
#include <xloil/StringUtils.h>
#include <xloil/Throw.h>
#include <xloil/WindowsSlim.h>
#include <xloil/Preprocessor.h>
#include <cstdlib>
#include <tomlplusplus/toml.hpp>

using std::vector;
using std::wstring;
using std::string;

namespace xloil
{
  namespace Python
  {
    static HMODULE thePythonLib = nullptr;
    static PluginInitFunc theInitFunc = nullptr;

    XLO_PLUGIN_INIT(AddinContext* context, const PluginContext& plugin)
    {
      try
      {
        if (plugin.action == PluginContext::Load)
        {
          linkPluginToCoreLogger(context, plugin);
          if (plugin.settings.empty())
            XLO_THROW(L"No settings found for {0} with addin {1}", plugin.pluginName, context->pathName());

          auto pyVer = utf8ToUtf16(plugin.settings["PythonVersion"].value_or(""));
          if (pyVer.empty())
            XLO_THROW("No xlOilPythonVersion specified in Python Environment block");

          // Convert X.Y version to XY and form the dll name
          auto dllName = fmt::format(L"xloil_Python{0}.dll", 
            pyVer.replace(pyVer.find(L'.'), 1, L""));

          // Load the library - the xlOil loader should already have set the DLL
          // load directory and we expect to find the versioned python plugins
          // in the directory this DLL is in.
          thePythonLib = LoadLibrary(dllName.c_str());
          if (!thePythonLib)
            XLO_THROW(L"Failed LoadLibrary for: {}", dllName);

          theInitFunc = (PluginInitFunc)GetProcAddress(thePythonLib,
            XLO_STR(XLO_PLUGIN_INIT_FUNC));
          if (!theInitFunc)
            XLO_THROW(L"Failed to find entry point {} in {}", 
              XLO_WSTR(XLO_PLUGIN_INIT_FUNC), dllName);
        }

        // Forward the request to the real python plugins 
        theInitFunc(context, plugin);
        return 0;
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(e.what());
        return -1;
      }
    }
  }
}