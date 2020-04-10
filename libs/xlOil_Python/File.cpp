#include "File.h"
#include "PyHelpers.h"
#include <xloil/Log.h>

using std::shared_ptr;
using std::vector;
using std::string;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    void scanModule(py::object& mod)
    {
      py::gil_scoped_acquire get_gil;

      auto oilModule = py::module::import("xloil");
      auto scanFunc = oilModule.attr("scan_module").cast<py::function>();

      try
      {
        XLO_INFO("Scanning module {0}", (string)py::str(mod));
        scanFunc.call(mod);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Error reading module {0}: {1}", (string)py::str(mod), e.what());
      }
    }
  }
}