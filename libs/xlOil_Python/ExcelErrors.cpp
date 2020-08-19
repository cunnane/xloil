#include "ExcelErrors.h"
#include "PyCoreModule.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;

namespace xloil {
  namespace Python {

    PyTypeObject* pyExcelErrorType = nullptr;

    namespace
    {
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind CellError type to xloil::CellError enum
        auto eType = py::enum_<CellError>(mod, "CellError");
        for (auto e : theCellErrors)
          eType.value(utf16ToUtf8(enumAsWCString(e)).c_str(), e);

        pyExcelErrorType = (PyTypeObject*)eType.ptr();
      });
    }
  }
}