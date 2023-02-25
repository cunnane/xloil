#include "PySource.h"

#include "PyHelpers.h"
#include "PyFunctionRegister.h"
#include "PyAddin.h"

#include <winreg/WinReg/WinReg.hpp>
#include <xlOil/AppObjects.h>
#include <xloil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelUI.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::vector;
using std::string;
using std::wstring;
using std::weak_ptr;
using std::wstring_view;
using winreg::RegKey;

namespace py = pybind11;

namespace
{
  /// <summary>
  /// Tries to find a local copy of a file given its OneDrive/Sharepoint URL.
  /// Uses the approach discussed here: https://stackoverflow.com/questions/33734706/
  /// </summary>
  bool oneDriveUrlToLocal(const wstring_view& url, wstring& path)
  {
    RegKey key;
    key.Open(HKEY_CURRENT_USER, L"Software\\SyncEngines\\Providers\\OneDrive");
    for (const auto& location : key.EnumSubKeys())
    {
      RegKey locationKey;
      locationKey.Open(key.Get(), location);

      auto urlNamespace = locationKey.GetStringValue(L"UrlNamespace");
      if (url.find(urlNamespace) != wstring_view::npos)
      {
        auto mountPoint = fs::path(locationKey.GetStringValue(L"MountPoint"));
        auto pathPart = url.substr(urlNamespace.size() + 2);
        path = mountPoint / pathPart;
        std::replace(path.begin(), path.end(), L'/', L'\\');
        while (!fs::exists(path))
        {
          auto nextSlash = pathPart.find_first_of(L'/');
          if (nextSlash == wstring_view::npos)
            return false;
          pathPart = pathPart.substr(nextSlash + 1);
          path = mountPoint / pathPart;
        }
        return true;
      }
    }
    return false;
  }

  /// <summary>
  /// Loads a text file directly from a URL. It does this via the Application.Open
  /// method so that if the URL is on OneDrive/Sharepoint Excel's own access tokens
  /// are leveraged.  Otherwise the user would need to get a Graph API token for xlOil,
  /// unless there is another way?  
  /// 
  /// Note this needs to be run on the main thread as it uses COM
  /// </summary>
  wstring loadOneDriveUrl(const wstring& url)
  {
    XLO_DEBUG(L"Loading module from OneDrive URL '{}'", url);
    auto wb = xloil::Application().open(url);
    auto firstSheet = wb.worksheets().list()[0];
    auto textRange = firstSheet.usedRange();
    wstring text;
    for (auto i = 0; i < textRange.nRows(); ++i)
    {
      auto value = textRange.value(i, 0);
      (text += value.toString()) += L"\n";
    }
    wb.close();

    return std::move(text);
  }
}

namespace xloil
{
  namespace Python
  {
    bool unloadModule(const py::handle& module)
    {
      py::gil_scoped_acquire get_gil;

      // Because xloil.scan_module adds workbook modules with the prefix
      // 'xloil.wb.', we can't simply lookup the module name in sys.modules.
      // We could rely on our knowledge of the prefix but iterating is not 
      // slow and is less fragile.
      auto sysModules = PyBorrow<py::dict>(PyImport_GetModuleDict());
      py::handle modName;
      for (auto [k, v] : sysModules)
        if (v.is(module))
          modName = k;

      if (!modName.ptr())
        return false;

      // Need to explictly clear the module's dict so that all globals get
      // dec-ref'd - they are not removed even when the module's ref-count 
      // hits zero.
      module.attr("__dict__").cast<py::dict>().clear();

      const auto ret = PyDict_DelItem(sysModules.ptr(), modName.ptr());

      // Remove last remaining reference to module
      module.dec_ref();

      return ret == 0;
    }

    namespace
    {
      std::map<wstring, wstring> theOneDriveSources;

      struct WorkbookOpenHandler
      {
        weak_ptr<PyAddin> _loadContext;

        WorkbookOpenHandler(const weak_ptr<PyAddin>& loadContext)
          : _loadContext(loadContext)
        {}

        void operator()(const wchar_t* wbPath, const wchar_t* wbName) const
        {
          auto addin = _loadContext.lock();
          const auto isUrl = wcsncmp(wbPath, L"http", 4) == 0;
          const auto separator = isUrl ? L'/' : L'\\';

          auto modulePath = formatStr(L"%s%c%s",
              wbPath,
              separator,
              addin->getLocalModulePath(wbName).c_str());

          XLO_DEBUG(L"Looking for workbook module at '{}'", modulePath);

          if (isUrl)
          {
            // TODO: better onedrive URL detection?
            bool isOneDrive = modulePath.find(L"sharepoint.com") != wstring::npos ||
              modulePath.find(L"docs.live.net") != wstring::npos;

            if (isOneDrive)
            {
              wstring localPath;
              if (oneDriveUrlToLocal(modulePath, localPath))
              {
                XLO_DEBUG(L"Found local copy of OneDrive file '{}' at '{}'", modulePath, localPath);
                modulePath = localPath;
              }
              else
              {
                theOneDriveSources[modulePath] = loadOneDriveUrl(modulePath);
              }
            }
          }
          else if (!fs::exists(modulePath))
            return;

          // First add the module, if the scan fails it will still be on the
          // file change watchlist. Note we always add workbook modules to the 
          // core context to avoid confusion.
          FunctionRegistry::addModule(_loadContext, modulePath, wbName);
          auto wbPathName = wstring(wbPath) + separator + wbName;

          py::gil_scoped_acquire getGil;
          addin->importFile(modulePath.c_str(), wbPathName.c_str());
        }
      };

      void checkExistingWorkbooks(const WorkbookOpenHandler& handler, Application& app)
      {
        for (const auto& wb : app.workbooks().list())
          handler(wb.path().c_str(), wb.name().c_str());
      }
    }

    std::shared_ptr<const void>
      createWorkbookOpenHandler(const weak_ptr<PyAddin>& loadContext, Application& app)
    {
      WorkbookOpenHandler handler(loadContext);

      checkExistingWorkbooks(handler, app);

      return Event::WorkbookOpen().bind(handler);
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
          mod.def("_get_onedrive_source", [](const wstring& url) { return theOneDriveSources[url]; });
      });
    }
  }
}