#include "PySource.h"

#include "PyHelpers.h"
#include "PyFunctionRegister.h"
#include "PyAddin.h"

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
        PyAddin& _loadContext;

        WorkbookOpenHandler(PyAddin& loadContext)
          : _loadContext(loadContext)
        {}

        void operator()(const wchar_t* wbPath, const wchar_t* wbName) const
        {
          auto modulePath = _loadContext.getLocalModulePath(
            fmt::format(L"{0}\\{1}", wbPath, wbName).c_str());

          std::error_code err;
          if (!fs::exists(modulePath, err))
            return;
          
          // First add the module, if the scan fails it will still be on the
          // file change watchlist. Note we always add workbook modules to the 
          // core context to avoid confusion.
          FunctionRegistry::addModule(_loadContext.context, modulePath, wbName);
          auto wbPathName = (fs::path(wbPath) / wbName).wstring();

          py::gil_scoped_acquire getGil;
          _loadContext.importFile(modulePath.c_str(), wbPathName.c_str());
        }
      };

      void checkExistingWorkbooks(const WorkbookOpenHandler& handler)
      {
        for (const auto& wb : App::Workbooks::list())
          handler(wb.path().c_str(), wb.name().c_str());
      }
    }
    std::shared_ptr<const void> 
      createWorkbookOpenHandler(PyAddin& loadContext)
    {
      WorkbookOpenHandler handler(loadContext);

      checkExistingWorkbooks(handler);

      return Event::WorkbookOpen().bind(handler);
    }
  }
}