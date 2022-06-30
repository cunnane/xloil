#include "BasicTypes.h"
#include "PyCore.h"
#include "PyEvents.h"
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil 
{
  namespace Python
  {
    namespace
    {
      struct CustomReturnConverter
      {
        std::shared_ptr<const IPyToExcel> value;
      };
      CustomReturnConverter* theConverter = nullptr;
    }

    const IPyToExcel* detail::getCustomReturnConverter()
    {
      return theConverter->value.get();
    }

    const char* IPyFromExcel::name() const
    {
      return "type";
    }

    namespace
    {
      template <class T>
      void convertPy(pybind11::module& mod, const char* type)
      {
        bindPyConverter<PyFromExcelConverter<T>>(mod, type).def(py::init<>());
      }

      template <class T>
      void convertXl(pybind11::module& mod, const char* type)
      {
        bindXlConverter<PyFuncToExcel<T>>(mod, type).def(py::init<>());
      }

      struct FromPyToCache
      {
        auto operator()(const PyObject* obj) const
        {
          return pyCacheAdd(PyBorrow<>(const_cast<PyObject*>(obj)));
        }
      };

      /// <summary>
      /// Always returns a single cell value. Uses the Excel object cache for 
      /// returned arrays and the Python object cache for unconvertable objects
      /// </summary>
      struct FromPyToSingleValue
      {
        auto operator()(const PyObject* obj) const
        {
          ExcelObj excelObj(FromPyObj()(obj));
          if (excelObj.isType(ExcelType::ArrayValue))
            return std::move(excelObj);
          return makeCached<ExcelObj>(std::move(excelObj));
        }
      };

      static int theBinder = addBinder([](py::module& mod)
      {
        // Bind converters for standard types
        convertPy<PyFromInt>(mod, "int");
        convertPy<PyFromDouble>(mod, "float");
        convertPy<PyFromBool>(mod, "bool");
        convertPy<PyFromString>(mod, "str");
        convertPy<PyFromAny>(mod, "object");
        convertPy<PyCacheObject>(mod, "Cache");

        convertPy<PyFromIntUncached>(mod, XLOPY_UNCACHED_PREFIX "int");
        convertPy<PyFromDoubleUncached>(mod, XLOPY_UNCACHED_PREFIX "float");
        convertPy<PyFromBoolUncached>(mod, XLOPY_UNCACHED_PREFIX "bool");
        convertPy<PyFromStringUncached>(mod, XLOPY_UNCACHED_PREFIX "str");
        convertPy<PyFromAnyUncached>(mod, XLOPY_UNCACHED_PREFIX "object");

        convertXl<FromPyLong>(mod, "int");
        convertXl<FromPyFloat>(mod, "float");
        convertXl<FromPyBool>(mod, "bool");
        convertXl<FromPyString>(mod, "str");
        convertXl<FromPyToCache>(mod, "Cache");
        convertXl<FromPyToSingleValue>(mod, "SingleValue");

        py::class_<CustomReturnConverter>(mod, "_CustomReturnConverter")
          .def_readwrite("value", &CustomReturnConverter::value);
          
        theConverter = new CustomReturnConverter();
        mod.add_object("return_converter",
          py::cast(theConverter, py::return_value_policy::take_ownership));
      });
    }
  }
}
