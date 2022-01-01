#include "PySource.h"

#include "PyHelpers.h"
#include "PyFunctionRegister.h"
#include "Main.h"
#include "EventLoop.h"

#include <xloil/Log.h>
#include <xlOil/ExcelThread.h>
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

    namespace
    {
      struct WorkbookOpenHandler
      {
        wstring _workbookPattern;
        PyAddin& _loadContext;

        WorkbookOpenHandler(const wstring& starredPattern, PyAddin& loadContext)
          : _loadContext(loadContext)
          , _workbookPattern(starredPattern)
        {
          // Turn the starred pattern into a fmt string for easier substitution later
          _workbookPattern.replace(_workbookPattern.find(L'*'), 1, wstring(L"{0}\\{1}"));
        }

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
          
          // First add the module, if the scan fails it will still be on the
          // file change watchlist. Note we always add workbook modules to the 
          // core context to avoid confusion.
          FunctionRegistry::addModule(_loadContext.context, modulePath, wbName);
          auto wbPathName = (fs::path(wbPath) / wbName).wstring();

          py::gil_scoped_acquire getGil;
          _loadContext.thread->callback("xloil.importer", "_import_file", 
            modulePath, _loadContext.pathName(), wbPathName);
        }
      };

      void checkExistingWorkbooks(const WorkbookOpenHandler& handler)
      {
        for (const auto& wb : App::Workbooks::list())
          handler(wb.path().c_str(), wb.name().c_str());
      }
    }
    std::shared_ptr<const void> 
      createWorkbookOpenHandler(const wchar_t* starredPattern, PyAddin& loadContext)
    {
      if (!wcschr(starredPattern, L'*'))
      {
        XLO_WARN("WorkbookModule should be of the form '*foo.py' where '*'"
          "will be replaced by the full workbook path with file extension removed");
        return std::shared_ptr<void>();
      }

      WorkbookOpenHandler handler(starredPattern, loadContext);

      checkExistingWorkbooks(handler);

      return Event::WorkbookOpen().bind(handler);
    }
  }
}