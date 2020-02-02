#include "InjectedModule.h"
#include "ExcelArray.h"
#include "BasicTypes.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil
{
  namespace Python
  {
    template <class TKeyConv, class TValConv>
    class PyDictFromArray : public ConverterImpl<PyObject*>
    {
      TKeyConv _keyConv;
      TValConv _valConv;
    public:
      PyObject * fromArray(const ExcelObj& obj) const
      {
        ExcelArray arr(obj, true);
        if (arr.nCols() != 2)
          XLO_THROW("Need a 2 column array to convert to dictionary");

        auto dict = py::dict();
        auto i = 0;

        for (; i < arr.nRows(); ++i)
        {
          auto key = PySteal<py::object>(_keyConv(arr(i, 0)));
          auto val = PySteal<py::object>(_valConv(arr(i, 1)));
          if (!key || !val || PyDict_SetItem(dict.ptr(), key.ptr(), val.ptr()) != 0)
            XLO_THROW("Failed to add row " + std::to_string(i) + " to dict");
        }

        return dict.release().ptr();
      }
    };

    PyObject* readKeywordArgs(const ExcelObj& obj)
    {
      return PyFromExcel<PyDictFromArray<FromExcel<PyFromString>, FromExcel<PyFromAny>>>()(obj);
    }

    namespace
    {
      template <class T>
      void declare(pybind11::module& mod, const char* name)
      {
        py::class_<T, IPyFromExcel, shared_ptr<T>>(mod, name)
          .def(py::init<>());
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
          declare<PyFromExcel<PyDictFromArray<FromExcel<PyFromAny>, FromExcel<PyFromAny>>>>(mod, "dict_object_from_Excel");
      });
    }
  }
}