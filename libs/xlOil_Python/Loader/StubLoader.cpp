#include <xloil/Plugin.h>
#include <xloil/StringUtils.h>
#include <xloil/Throw.h>
#include <xloil/WindowsSlim.h>
#include <cstdlib>

#define Py_LIMITED_API
#include <Python.h>

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
          throwIfNotExactVersion(plugin);

          // This means we need to link python3.dll. If python 4 comes along and 
          // the old python3.dll is no longer around, we could sniff the dependencies
          // of python.exe in PYTHONHOME to work out which version we need to load
          auto pyVersion = Py_GetVersion();

          string dllName = "xlOil_Python";

          // Version string looks like "X.Y.Z blahblahblah"
          // Convert X.Y version to XY and append to dllname. Stop processing
          // when we hit something else
          auto periods = 0;
          for (auto c = pyVersion; ;++c)
          {
            if (isdigit(*c))
              dllName.push_back(*c);
            else if (*c == '.')
            {
              if (++periods > 1)
                break;
            }
            else
              break;
          }
          dllName += ".dll";
         
          // Load the library - the xlOil loader should already have set the DLL
          // load directory and we expect to find the versioned python plugins
          // in the directory this DLL is in.
          thePythonLib = LoadLibraryA(dllName.c_str());
          if (!thePythonLib)
            XLO_THROW("Failed LoadLibrary for: {}", dllName);

          theInitFunc = (PluginInitFunc)GetProcAddress(
            thePythonLib,
            "xloil_python_init");

          if (!theInitFunc)
            XLO_THROW("Failed to find xloil python entry point in {}", dllName);
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