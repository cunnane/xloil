#include <xloil/Interface.h>
#include <xloilHelpers/StringUtils.h>
#include <xloil/Throw.h>
#include <xloilHelpers/WindowsSlim.h>
#include <cstdlib>
#include <tomlplusplus/toml.hpp>
#include <boost/preprocessor/stringize.hpp>

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
          if (!plugin.settings)
            XLO_THROW(L"No settings found for {0} with addin {1}", plugin.pluginName, context->pathName());

          auto pyEnv = (*plugin.settings)["Environment"].as_array();
          string pyVer;
          if (pyEnv)
            for (auto& table : *pyEnv)
              if (table.as_table()->contains("xlOilPythonVersion"))
              {
                pyVer = table.as_table()->get("xlOilPythonVersion")->value_or("");
                break;
              }

          if (pyVer.empty())
            XLO_THROW("No xlOilPythonVersion specified in Python Environment block");

          // Convert X.Y version to XY and form the dll name
          auto dllName = fmt::format("xloil_Python{0}.dll", 
            pyVer.replace(pyVer.find('.'), 1, ""));

          // Load the library - the xlOil loader should already have set the DLL
          // load directory and we expect to find the versioned python plugins
          // in the directory this DLL is in.
          thePythonLib = LoadLibrary(dllName.c_str());
          if (!thePythonLib)
            return -1;

          theInitFunc = (PluginInitFunc)GetProcAddress(thePythonLib,
            BOOST_PP_STRINGIZE(XLO_PLUGIN_INIT_FUNC));
          if (!theInitFunc)
            return -1;
        }

        // Forward the request to the real python plugins 
        theInitFunc(context, plugin);
        return 0;
      }
      catch (...)
      {
        return -1;
      }
    }
  }
}