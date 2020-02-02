#include "BasicTypes.h"
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil 
{
  namespace Python
  {
    PyTypeObject* pyExcelErrorType = nullptr;

    template <class T>
    void declare(pybind11::module& mod, const char* type)
    {
      bindFrom<T>(mod, type).def(py::init<>());
    }

    static int theBinder = addBinder([](py::module& mod)
    {

     // std::is_base_of< IPyFromExcel, PyFromExcel<PyFromInt>>::value;
      //template <typename Base, typename Derived> using is_strict_base_of = bool_constant<
      //std::is_base_of<Base, Derived>::value && !std::is_same<Base, Derived>::value > ;
      //auto thing = py::detail::is_strict_base_of<IPyFromExcel, PyFromExcel<PyFromInt>>::value;
      py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel");
      py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

      auto eType = py::enum_<CellError>(mod, "CellError");
      for (auto e : theCellErrors)
        eType.value(wstring_to_utf8(toWCString(e)).c_str(), e);

      pyExcelErrorType = (PyTypeObject*)eType.get_type().ptr();

      declare<PyFromExcel<PyFromInt>>(mod, "int");
      declare<PyFromExcel<PyFromDouble>>(mod, "float");
      declare<PyFromExcel<PyFromBool>>(mod, "bool");
      declare<PyFromExcel<PyFromString>>(mod, "str");
      declare<PyFromExcel<PyFromAny>>(mod, "object");
      declare<PyFromExcel<PyCacheObject>>(mod, "cached");
    });

    PyObject * PyFromAny::fromError(CellError err) const
    {
      auto pyObj = py::cast(err);
      return pyObj.release().ptr();
    }
  }
}
