#include "PyCoreModule.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil
{
  namespace Python
  {
    template <class TKeyConv, class TValConv>
    class PyDictFromArray : public PyFromCache<PyDictFromArray<TKeyConv, TValConv>>
    {
      TKeyConv _keyConv;
      TValConv _valConv;
    public:
      PyObject* fromArray(const ExcelObj& obj) const
      {
        return fromArrayObj(ExcelArray(obj, true));
      }
      PyObject* fromArrayObj(const ExcelArray& arr) const
      {
        if (arr.nCols() != 2)
          XLO_THROW("Need a 2 column array to convert to dictionary");

        auto dict = py::dict();
        ExcelArray::row_t i = 0;

        for (; i < arr.nRows(); ++i)
        {
          auto key = PySteal<py::object>(_keyConv(arr.at(i, 0)));
          auto val = PySteal<py::object>(_valConv(arr.at(i, 1)));
          if (!key || !val || PyDict_SetItem(dict.ptr(), key.ptr(), val.ptr()) != 0)
            XLO_THROW("Failed to add row " + std::to_string(i) + " to dict");
        }

        return dict.release().ptr();
      }
    };

    PyObject* readKeywordArgs(const ExcelObj& obj)
    {
      return PyFromExcel<PyDictFromArray<FromExcel<PyFromString>, FromExcel<PyFromAny<>>>>()(obj);
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
          declare<PyFromExcel<PyDictFromArray<PyFromExcel<PyFromAny<>>, PyFromExcel<PyFromAny<>>>>>(mod, "dict_object_from_Excel");
      });
    }
  }
}