#include "PyCoreModule.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "ArrayHelpers.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;
using std::shared_ptr;

namespace xloil
{
  namespace Python
  {
    namespace detail
    {
      template <class TKeyConv, class TValConv>
      class PyDictFromArray : public PyFromExcelImpl
      {
        TKeyConv _keyConv;
        TValConv _valConv;
      public:
        using PyFromExcelImpl::operator();

        // Return empty dictionary for missing value
        PyObject* operator()(MissingVal) const
        {
          return PyDict_New();
        }
        // Return empty dictionary for nil value
        PyObject* operator()(nullptr_t) const
        {
          return PyDict_New();
        }

        PyObject* operator()(ArrayVal obj) const
        {
          ExcelArray arr(obj);

          // Return empty dict is array is empty
          if (arr.nRows() == 0)
            return PyDict_New();

          if (arr.nCols() != 2)
            XLO_THROW("Need a 2 column array to convert to dictionary");

          auto dict = py::dict();
          ExcelArray::row_t i = 0;

          for (; i < arr.nRows(); ++i)
          {
            auto key = PySteal(_keyConv(arr.at(i, 0)));
            auto val = PySteal(_valConv(arr.at(i, 1)));
            if (!key || !val || PyDict_SetItem(dict.ptr(), key.ptr(), val.ptr()) != 0)
              XLO_THROW("Failed to add row " + std::to_string(i) + " to dict");
          }

          return dict.release().ptr();
        }
        constexpr wchar_t* failMessage() const { return L"Expected array"; }
      };
    }
    using PyDictFromExcel = PyFromExcelConverter<
      detail::PyDictFromArray<PyFromAny, PyFromAny>>;

    PyObject* readKeywordArgs(const ExcelObj& obj)
    {
      return PyFromExcel<
        detail::PyDictFromArray<PyFromString, PyFromAny>>()(obj);
    }
   
    class XlFromDict: public IPyToExcel
    {
    public:
      ExcelObj operator()(const PyObject& obj) const override
      {
        auto p = (PyObject*)&obj;
        if (!PyDict_Check(p))
          return ExcelObj();
        const auto size = PyDict_Size(p);

        size_t stringLength = 0;
        PyObject *key, *value;
        Py_ssize_t pos = 0;

        while (PyDict_Next(p, &pos, &key, &value)) 
        {
          accumulateObjectStringLength(key, stringLength);
          accumulateObjectStringLength(value, stringLength);
        }

        ExcelArrayBuilder builder((ExcelObj::row_t)size, 2, stringLength);

        pos = 0;
        ExcelObj::row_t row = 0; // Cannot use pos - it is an internal pointer only
        while (PyDict_Next(p, &pos, &key, &value))
        {
          builder(row, 0).emplace(FromPyObj()(key, builder.charAllocator()));
          builder(row, 1).emplace(FromPyObj()(value, builder.charAllocator()));
          ++row;
        }

        return builder.toExcelObj();
      }
    };

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
          bindPyConverter<PyDictFromExcel>(mod, "dict").def(py::init<>());
          bindXlConverter<XlFromDict>(mod, "dict").def(py::init<>());
      });
    }
  }
}