
#include <xloil/Plugin.h>
#include <xloil/Log.h>

#include <xloil/Preprocessor.h>
#include <tomlplusplus/toml.hpp>
#include <xloil/WindowsSlim.h>

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
          if (plugin.failIfNotExactVersion())
            return -1;
          wstring dllName, pyVer;
          //linkPluginToCoreLogger(context, plugin);
          if (plugin.settings.empty())
            plugin.error(L"No settings found for {0} with addin {1}", plugin.pluginName, context->pathName());

          pyVer = utf8ToUtf16(plugin.settings["PythonVersion"].value_or(""));
          if (pyVer.empty())
            plugin.error(L"No xlOilPythonVersion specified in Python Environment block");

          // Convert X.Y version to XY and form the dll name

          
          dllName = XLO_FMT(L"xloil_Python{0}.dll", 
            pyVer.replace(pyVer.find(L'.'), 1, L""));

          // Load the library - the xlOil loader should already have set the DLL
          // load directory and we expect to find the versioned python plugins
          // in the directory this DLL is in.
          //thePythonLib = LoadLibrary(dllName.c_str());
          //if (!thePythonLib)
          //  plugin.error(L"Failed LoadLibrary for: {}", dllName);

          //theInitFunc = (PluginInitFunc)GetProcAddress(thePythonLib,
          //  XLO_STR(XLO_PLUGIN_INIT_FUNC));
          //if (!theInitFunc)
          //  plugin.error(L"Failed to find entry point {} in {}",
          //    XLO_WSTR(XLO_PLUGIN_INIT_FUNC), dllName);
        }

        // Forward the request to the real python plugins 
        return theInitFunc(context, plugin);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(e.what());
        return -1;
      }
    }
  }
}