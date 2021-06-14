#include "BasicTypes.h"
#include "PyCoreModule.h"
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil 
{
  namespace Python
  {
    template <class T>
    void convertPy(pybind11::module& mod, const char* type)
    {
      bindPyConverter<PyExcelConverter<T>>(mod, type).def(py::init<>());
    }
    
    template<class TFunc>
    class PyFuncToXl : public IPyToExcel
    {
    public:
      ExcelObj operator()(const PyObject& obj) const override
      {
        return TFunc()(&obj);
      }
    };

    template <class T>
    void convertXl(pybind11::module& mod, const char* type)
    {
      bindXlConverter<PyFuncToXl<T>>(mod, type).def(py::init<>());
    }

    shared_ptr<const IPyToExcel> theCustomReturnConverter = nullptr;

    void setReturnConverter(shared_ptr<const IPyToExcel> conv)
    {
      theCustomReturnConverter = conv;
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

      convertXl<FromPyLong>(mod, "int");
      convertXl<FromPyFloat>(mod, "float");
      convertXl<FromPyBool>(mod, "bool");
      convertXl<FromPyString>(mod, "str");
      convertXl<FromPyToCache>(mod, "Cache");
      convertXl<FromPyToSingleValue>(mod, "SingleValue");

      mod.def("set_return_converter", setReturnConverter);
    });
  }
}
