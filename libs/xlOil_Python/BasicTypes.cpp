#include "BasicTypes.h"

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil 
{
  namespace Python
  {
    template <class T>
    void declare(pybind11::module& mod, const char* type)
    {
      bindFrom<T>(mod, type).def(py::init<>());
    }
    
    static int theBinder = addBinder([](py::module& mod)
    {
      // Bind converters for standard types
      declare<PyFromExcel<PyFromInt>>(mod, "int");
      declare<PyFromExcel<PyFromDouble>>(mod, "float");
      declare<PyFromExcel<PyFromBool>>(mod, "bool");
      declare<PyFromExcel<PyFromString>>(mod, "str");
      declare<PyFromExcel<PyFromAny<>>>(mod, "object");
      declare<PyFromExcel<PyCacheObject>>(mod, "cache");
    });
  }
}
