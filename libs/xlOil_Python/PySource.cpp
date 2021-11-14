#include "PySource.h"

#include "PyHelpers.h"
#include "PyFunctionRegister.h"
#include "Main.h"

#include <xloil/Log.h>
#include <xlOil/ExcelApp.h>
#include <xlOil/Events.h>
#include <xlOil/ExcelUI.h>
#include <filesystem>

namespace fs = std::filesystem;

using std::vector;
using std::string;
using std::wstring;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    void scanModule(const py::object& mod)
    {
      py::gil_scoped_acquire getGil;

      const auto xloilModule = py::module::import("xloil");
      const auto scanFunc = xloilModule.attr("scan_module").cast<py::function>();

      const auto modName = (string)py::str(mod);
      try
      {
        XLO_INFO("Scanning module {0}", modName);
        scanFunc(mod);
      }
      catch (const std::exception& e)
      {
        auto pyPath = (string)py::str(PyBorrow<py::list>(PySys_GetObject("path")));
        XLO_ERROR("Error reading module {0}: {1}\nsys.path={2}", 
          modName, e.what(), pyPath);
      }
    }
    py::object loadModuleFromFile(const wchar_t* filepath, const wchar_t* linkedWorkbook)
    {
      py::gil_scoped_acquire getGil;
      // TODO: make this into a global
      const auto xloilModule = py::module::import("xloil.importer");
      return xloilModule.attr("import_from_file").cast<py::function>()(py::wstr(filepath),
        linkedWorkbook ? py::wstr(linkedWorkbook) : py::str());
    }
    bool unloadModule(const py::handle& module)
    {
      py::gil_scoped_acquire get_gil;

      // Because xloil.scan_module adds workbook modules with the prefix
      // 'xloil.wb.', we can't simply lookup the module name in sys.modules.
      // We could rely on our knowledge of the prefix but iterating is not 
      // slow and is less fragile.
      auto sysModules = PyBorrow<py::dict>(PyImport_GetModuleDict());
      py::handle modName;
      for (auto[k, v] : sysModules)
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

    struct WorkbookOpenHandler
    {
      WorkbookOpenHandler(const wstring& starredPattern)
      {
        // Turn the starred pattern into a fmt string for easier substitution later
        _workbookPattern = starredPattern;
        _workbookPattern.replace(_workbookPattern.find(L'*'), 1, wstring(L"{0}\\{1}"));
      }

      wstring _workbookPattern;

      void operator()(const wchar_t* wbPath, const wchar_t* wbName) const
      {
        // Subtitute in to find target module name, removing extension
        auto fileExtn = wcsrchr(wbName, L'.');
        auto modulePath = fmt::format(_workbookPattern,
          wbPath,
          fileExtn ? wstring(wbName, fileExtn).c_str() : wbName);

        std::error_code err;
        if (!fs::exists(modulePath, err))
          return;

        try
        {
          // First add the module, if the scan fails it will still be on the
          // file change watchlist. Note we always add workbook modules to the 
          // core context to avoid confusion.
          FunctionRegistry::addModule(theCoreContext, modulePath, wbName);
          runPython([modulePath, workbook = wstring(wbName)](int)
          {
            StatusBar status(5000);
            status.msg(L"Loading " + workbook + L"...");
            auto loaded = loadModuleFromFile(modulePath.c_str(), workbook.c_str());
            scanModule(loaded);
            status.msg(L"Finished loading " + workbook);
          });
        }
        catch (const std::exception& e)
        {
          XLO_ERROR(L"Failed to load module {0}: {1}", modulePath, utf8ToUtf16(e.what()));
        }
      }
    };

    void checkWorkbooksOnOpen(const WorkbookOpenHandler& handler)
    {
      for (const auto& wb : App::workbooks())
        handler(wb.path().c_str(), wb.name().c_str());
    }

    void createWorkbookOpenHandler(const wchar_t* starredPattern)
    {
      if (!wcschr(starredPattern, L'*'))
      {
        XLO_WARN("WorkbookModule should be of the form '*foo.py' where '*'"
          "will be replaced by the full workbook path with file extension removed");
        return;
      }
      WorkbookOpenHandler handler(starredPattern);

      checkWorkbooksOnOpen(handler);

      static auto wbOpenHandler = Event::WorkbookOpen().bind(handler);
    }
  }
}