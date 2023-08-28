#include "PyTupleType.h"
#include "PyCore.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "ArrayHelpers.h"
#include <xlOil/ArrayBuilder.h>
#include <pybind11/pybind11.h>


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
          while ((innerItem = PyIter_Next(innerIter)) != nullptr)
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

      if (nRows > XL_MAX_ROWS)
        XLO_THROW("Max rows exceeded when returning iterator");
      if (nCols > XL_MAX_COLS)
        XLO_THROW("Max columns exceeded when returning iterator");


      if (PyErr_Occurred())
        throw py::error_already_set();

      Py_DECREF(iter);

      // Python supports an empty tuple, but Excel doesn't support an
      // empty array, so return a Missing type
      if (nRows == 0)
        return ExcelObj(ExcelType::Missing);

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
            builder(i, j++).take(FromPyObj()(innerItem, builder.charAllocator()));
            Py_DECREF(innerItem);
          }
          if (PyErr_Occurred())
            throw py::error_already_set();
          Py_DECREF(innerIter);
        }
        else
          builder(i, j++).take(FromPyObj()(item, builder.charAllocator()));

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
    class PyTupleFromArray : public detail::PyFromExcelImpl
    {
      TValConv _valConv;
    public:
      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "tuple";

      PyObject* operator()(const ArrayVal& obj)
      {
        ExcelArray arr(obj);
        auto nRows = arr.nRows();
        auto nCols = arr.nCols();

        auto outer = py::tuple(nRows);
        for (decltype(nRows) i = 0; i < nRows; ++i)
        {
          auto inner = py::tuple(nCols);
          for (decltype(nCols) j = 0; j < nCols; ++j)
          {
            auto val = _valConv(arr.at(i, j));
            PyTuple_SET_ITEM(inner.ptr(), j, val);
          }
          PyTuple_SET_ITEM(outer.ptr(), i, inner.release().ptr());
        }
        return outer.release().ptr();
      }
      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    template <class TValConv>
    class PyListFromArray : public detail::PyFromExcelImpl
    {
      TValConv _valConv;
    public:
      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "list";

      PyObject* operator()(const ArrayVal& obj)
      {
        ExcelArray arr(obj);
        auto nRows = arr.nRows();
        auto nCols = arr.nCols();

        auto outer = py::list(nRows);
        for (decltype(nRows) i = 0; i < nRows; ++i)
        {
          auto inner = py::list(nCols);
          for (decltype(nCols) j = 0; j < nCols; ++j)
          {
            auto val = _valConv(arr.at(i, j));
            PyList_SET_ITEM(inner.ptr(), j, val);
          }
          PyList_SET_ITEM(outer.ptr(), i, inner.release().ptr());
        }
        return outer.release().ptr();
      }
      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    PyObject* excelArrayToNestedTuple(const ExcelObj & obj)
    {
      return PyTupleFromArray<PyFromAny>()(static_cast<const ArrayVal&>(obj));
    }

    namespace
    {
      struct Adapter
      {
        template<class... Args> auto operator()(Args&&... args)
        {
          return nestedIterableToExcel(std::forward<Args>(args)...);
        }
        static constexpr char* ourName = "iterable";
      };
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        bindPyConverter<PyFromExcelConverter<PyTupleFromArray<PyFromAny>>>(mod, "tuple").def(py::init<>());
        bindPyConverter<PyFromExcelConverter<PyListFromArray<PyFromAny>>>(mod, "list").def(py::init<>());
        auto tupleToExcel = bindXlConverter<PyFuncToExcel<Adapter>>(mod, "tuple").def(py::init<>());
        mod.add_object((std::string(theReturnConverterPrefix) + "list").c_str(), tupleToExcel);
        mod.add_object((std::string(theReturnConverterPrefix) + "iterable").c_str(), tupleToExcel);
      });
    }
  }
}