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

      auto iter = py::iter(py::handle(p));

      size_t stringLength = 0;
      uint32_t nRows = 0;
      uint16_t nCols = 1;

      // First loop to establish array size and length of strings
      while (iter != py::iterator::sentinel())
      {
        if (PyIterable_Check(iter->ptr()) && !PyUnicode_Check(iter->ptr()))
        {
          decltype(nCols) j = 0;
          auto innerIter = py::iter(*iter);
          while (innerIter != py::iterator::sentinel())
          {
            ++j;
            accumulateObjectStringLength(innerIter->ptr(), stringLength);
            ++innerIter;
          }
          
          nCols = std::max(nCols, j);
        }
        else
          accumulateObjectStringLength(iter->ptr(), stringLength);

        ++nRows;
        ++iter;
      }

      if (nRows > XL_MAX_ROWS)
        XLO_THROW("Max rows exceeded when returning iterator");
      if (nCols > XL_MAX_COLS)
        XLO_THROW("Max columns exceeded when returning iterator");

      if (PyErr_Occurred())
        throw py::error_already_set();

      // Python supports an empty tuple, but Excel doesn't support an
      // empty array, so return a Missing type
      if (nRows == 0)
        return ExcelObj(ExcelType::Missing);

      ExcelArrayBuilder builder(nRows, nCols, stringLength);

      // Second loop to fill in array values
      iter = py::iter(py::handle(p));
      size_t i = 0, j = 0;
      while (iter != py::iterator::sentinel())
      {
        j = 0;
        if (PyIterable_Check(iter->ptr()) && !PyUnicode_Check(iter->ptr()))
        {
          auto innerIter = py::iter(*iter);
          while (innerIter != py::iterator::sentinel())
          {
            builder(i, j++).take(FromPyObj<detail::ReturnToCache, true>()(innerIter->ptr(), builder.charAllocator()));
          }
        }
        else
          builder(i, j++).take(FromPyObj<detail::ReturnToCache, true>()(iter->ptr(), builder.charAllocator()));

        if (PyErr_Occurred())
          throw py::error_already_set();

        // Fill with N/A
        for (; j < nCols; ++j)
          builder(i, j) = CellError::NA;

        ++i;
        ++iter;
      }

      return builder.toExcelObj();
    }

    template <class TValConv>
    class PyTupleFromArray : public detail::PyFromExcelImpl<PyTupleFromArray<TValConv>>
    {
      TValConv _valConv;
    public:
      using detail::PyFromExcelImpl<PyTupleFromArray<TValConv>>::operator();
      static constexpr char* const ourName = "tuple";

      PyObject* operator()(const ExcelObj& obj) const
      {
        ExcelArray arr(cacheCheck(obj));
        if (arr.dims() < 2)
        {
          auto result = py::tuple(arr.size());
          for (auto i = 0u; i < arr.size(); ++i)
          {
            auto val = _valConv(arr.at(i));
            PyTuple_SET_ITEM(result.ptr(), i, val);
          }
          return result.release().ptr();
        }
        else
        {
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
      }
      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    template <class TValConv>
    class PyListFromArray : public detail::PyFromExcelImpl<PyListFromArray<TValConv>>
    {
      TValConv _valConv;
    public:
      using detail::PyFromExcelImpl<PyListFromArray<TValConv>>::operator();
      static constexpr char* const ourName = "list";

      PyObject* operator()(const ExcelObj& obj) const
      {
        ExcelArray arr(cacheCheck(obj));

        if (arr.dims() < 2)
        {
          auto result = py::list(arr.size());
          for (auto i = 0u; i < arr.size(); ++i)
          {
            auto val = _valConv(arr.at(i));
            PyList_SET_ITEM(result.ptr(), i, val);
          }
          return result.release().ptr();
        }
        else
        {
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
      }
      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    PyObject* excelArrayToNestedTuple(const ExcelObj & obj)
    {
      return PyTupleFromArray<PyFromAny>()(obj);
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