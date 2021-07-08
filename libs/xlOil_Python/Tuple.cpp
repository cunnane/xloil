#include "PyCoreModule.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "ArrayHelpers.h"
#include <xlOil/ArrayBuilder.h>
#include <pybind11/pybind11.h>
#include "Tuple.h"

namespace py = pybind11;
using std::shared_ptr;

namespace xloil
{
  namespace Python
  {
    ExcelObj nestedIterableToExcel(const PyObject* obj)
    {
      auto p = const_cast<PyObject*>(obj);

      assert(PyIterable_Check(p));

      auto* iter = PyObject_GetIter(p);
      if (!iter)
        XLO_THROW("nestedIterableToExcel: could not create iterator");

      size_t stringLength = 0;
      uint32_t nRows = 0;
      uint16_t nCols = 1;
      PyObject *item, *innerItem;

      // First loop to establish array size and length of strings
      while ((item = PyIter_Next(iter)) != 0) 
      {
        ++nRows;
        if (PyIterable_Check(item) && !PyUnicode_Check(item))
        {
          decltype(nCols) j = 0;
          auto* innerIter = PyCheck(PyObject_GetIter(item));
          while ((innerItem = PyIter_Next(innerIter)) != 0)
          {
            ++j;
            accumulateObjectStringLength(innerItem, stringLength);
            Py_DECREF(innerItem);
          }
          Py_DECREF(innerIter);

          if (PyErr_Occurred())
            throw py::error_already_set();

          nCols = std::max(nCols, j);
        }
        else
          accumulateObjectStringLength(item, stringLength);
        Py_DECREF(item);
      }

      if (PyErr_Occurred())
        throw py::error_already_set();

      Py_DECREF(iter);

      ExcelArrayBuilder builder(nRows, nCols, stringLength);

      // Second loop to fill in array values
      iter = PyObject_GetIter(p);
      size_t i = 0, j = 0;
      while ((item = PyIter_Next(iter)) != 0)
      {
        j = 0;
        if (PyIterable_Check(item) && !PyUnicode_Check(item))
        {
          auto* innerIter = PyCheck(PyObject_GetIter(item));
          while ((innerItem = PyIter_Next(innerIter)) != 0)
          {
            builder(i, j++).emplace(FromPyObj()(innerItem, builder.charAllocator()));
            Py_DECREF(innerItem);
          }
          if (PyErr_Occurred())
            throw py::error_already_set();
          Py_DECREF(innerIter);
        }
        else
          builder(i, j++).emplace(FromPyObj()(item, builder.charAllocator()));

        // Fill with N/A
        for (; j < nCols; ++j)
          builder(i, j) = CellError::NA;

        Py_DECREF(item);
        ++i;
      }
      Py_DECREF(iter);

      if (PyErr_Occurred())
        throw py::error_already_set();

      return builder.toExcelObj();
    }

    template <class TValConv>
    class PyTupleFromArray : public PyFromExcelImpl
    {
      TValConv _valConv;
    public:
      using PyFromExcelImpl::operator();

      PyObject* operator()(ArrayVal obj) const
      {
        ExcelArray arr(obj);
        auto nRows = arr.nRows();
        auto nCols = arr.nCols();

        auto outer = py::tuple(nRows);
        for (decltype(nRows) i = 0; i < nRows; ++i)
        {
          auto inner = py::tuple(nCols);
          PyTuple_SET_ITEM(outer.ptr(), i, inner.ptr());
          for (decltype(nCols) j = 0; j < nCols; ++j)
          {
            auto val = _valConv(arr.at(i, j));
            PyTuple_SET_ITEM(inner.ptr(), j, val);
          }
        }
        return outer.release().ptr();
      }
      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    PyObject* excelArrayToNestedTuple(const ExcelObj & obj)
    {
      return PyTupleFromArray<PyFromAny>()(ArrayVal{ obj });
    }

    namespace
    {
      struct Adapter
      {
        template<class... Args> auto operator()(Args&&... args)
        {
          return nestedIterableToExcel(std::forward<Args>(args)...);
        }
      };
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        bindPyConverter<PyFromExcelConverter<PyTupleFromArray<PyFromAny>>>(mod, "tuple_from_Excel").def(py::init<>());
        bindXlConverter<PyFuncToExcel<Adapter>>(mod, "tuple_to_Excel").def(py::init<>());
      });
    }
  }
}