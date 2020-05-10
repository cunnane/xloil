#include "File.h"
#include "PyHelpers.h"
#include <xloil/Log.h>

using std::vector;
using std::string;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    void scanModule(const py::object& mod, const wchar_t* workbookName)
    {
      py::gil_scoped_acquire get_gil;

      const auto xloilModule = py::module::import("xloil");
      const auto scanFunc = xloilModule.attr("scan_module").cast<py::function>();

      const auto wbName = workbookName 
        ? py::wstr(workbookName) 
        : py::none();

      // Any need for this?
      //if (PyModule_Check(mod.ptr())) Event_PyReload().fire(mod);

      const auto modName = (string)py::str(mod);
      try
      {
        XLO_INFO("Scanning module {0}", modName);
        scanFunc.call(mod, wbName);
      }
      catch (const std::exception& e)
      {
        auto pyPath = (string)py::str(PyBorrow<py::list>(PySys_GetObject("path")));
        XLO_ERROR("Error reading module {0}: {1}\nsys.path={2}", 
          modName, e.what(), pyPath);
      }
    }
  }
}