#include <xloil/Plugin.h>
#include <xloil/StringUtils.h>
#include <xlOilHelpers/Environment.h>
#include <xloil/Throw.h>
#include <xloil/WindowsSlim.h>
#include <cstdlib>
#include <filesystem>
#include <fstream>

namespace fs = std::filesystem;
using std::vector;
using std::wstring;
using std::string;

namespace
{
  /// <summary>
  /// Finds the value of the home= directive in a pyvenv.cfg.
  /// Verbatim from CPython's launcher.c
  /// </summary>
  int find_home_value(const char* buffer, const char** start, size_t* length)
  {
    for (const char* s = strstr(buffer, "home"); s; s = strstr(s + 1, "\nhome")) {
      if (*s == '\n') {
        ++s;
      }
      for (auto i = 4; i > 0 && *s; --i, ++s);

      while (*s && iswspace(*s)) {
        ++s;
      }
      if (*s != L'=') {
        continue;
      }

      do {
        ++s;
      } while (*s && iswspace(*s));

      *start = s;
      auto* nl = strchr(s, '\n');
      if (nl) {
        *length = (ptrdiff_t)nl - (ptrdiff_t)s;
      }
      else {
        *length = strlen(s);
      }
      return 1;
    }
    return 0;
  }
}
namespace xloil
{
  namespace Python
  {
    namespace
    {
      std::string findPythonHomeDir()
      {
        constexpr auto PYVENVCFG = "pyvenv.cfg";

        auto pyExecutableDir = fs::path(getEnvironmentVar("PYTHONEXECUTABLE")).parent_path();
        auto pyvenvCfg = pyExecutableDir / PYVENVCFG;

        if (!fs::exists(pyvenvCfg))
        {
          pyvenvCfg = pyExecutableDir.parent_path() / PYVENVCFG;
          if (!fs::exists(pyvenvCfg))
          {
            // Not a venv
            return pyExecutableDir.string();
          }
        }

        std::ifstream file(pyvenvCfg);
        std::string contents(std::istreambuf_iterator<char>{file}, {});
        const char* pHomeDir;
        size_t homeDirLength;
        if (find_home_value(contents.c_str(), &pHomeDir, &homeDirLength) == 1)
          return std::string(pHomeDir, homeDirLength);

        else
          XLO_THROW("Could not find 'home' in {0}", pyvenvCfg.string());
      }

      static HMODULE thePythonLib = nullptr;
      static PluginInitFunc theInitFunc = nullptr;
    }

    XLO_PLUGIN_INIT(AddinContext* context, const PluginContext& plugin)
    {
      try
      {
        if (plugin.action == PluginContext::Load)
        {
          linkPluginToCoreLogger(context, plugin);
          throwIfNotExactVersion(plugin);

          string dllName = "xlOil_Python";

          auto userVersion = getEnvironmentVar("XLOIL_PYTHON_VERSION");

          auto pythonHome = findPythonHomeDir();
          auto path = getEnvironmentVar("PATH")
            .append(";").append(pythonHome)
            .append(";").append(pythonHome).append("\\Library\\bin");
          setEnvironmentVar("PATH", path.c_str());

          if (!userVersion.empty())
          {
            userVersion.erase(
              std::remove(userVersion.begin(), userVersion.end(), '.'), 
              userVersion.end());
            dllName += userVersion;
          }
          else
          {
            // Load the plugin, but first set the DLL directory to our PYTHONHOME in case
            // there are other pythons on the path
            HMODULE py3dll;
            {
              PushDllDirectory setDllDir(pythonHome.c_str());
              // Load the plugin
              py3dll = LoadLibrary(L"python3.dll");
              if (!py3dll)
                XLO_THROW(writeWindowsError());
            }

            typedef const char*(*Py_GetVersion_t)();
            const auto pyGetVersion = (Py_GetVersion_t)
              GetProcAddress(py3dll, "Py_GetVersion");

            if (!pyGetVersion)
              XLO_THROW(writeWindowsError());

            // This means we need to link python3.dll. If python 4 comes along and 
            // the old python3.dll is no longer around, we could sniff the dependencies
            // of python.exe to work out which version we need to load
            auto pyVersion = pyGetVersion();

            // Version string looks like "X.Y.Z blahblahblah"
            // Convert X.Y version to XY and append to dllname. Stop processing
            // when we hit something else
            auto periods = 0;
            for (auto c = pyVersion; ; ++c)
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
          }

          dllName += ".pyd";

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