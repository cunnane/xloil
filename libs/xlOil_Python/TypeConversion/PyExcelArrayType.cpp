#pragma once
#include "PyHelpers.h"
#include "PyExcelArrayType.h"
#include "PyCore.h"
#include "BasicTypes.h"

using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    extern PyTypeObject* ExcelArrayType = nullptr;

    /// <summary>
    /// This pretty weird looking ctor is supports adding to the refcount
    /// when taking subarrays via the [] operator
    /// </summary>
    PyExcelArray::PyExcelArray(
      const PyExcelArray& from,
      ExcelArray&& rebase)
      : _base(std::move(rebase))
      , _refCount(from._refCount)
    {
      *_refCount += 1;
    }

    PyExcelArray::PyExcelArray(const PyExcelArray& from)
      : _base(from._base)
      , _refCount(from._refCount)
    {
      *_refCount += 1;
    }

    PyExcelArray::PyExcelArray(ExcelArray&& arr)
      : _base(std::move(arr))
      , _refCount(new size_t(1))
    {}

    PyExcelArray::PyExcelArray(const ExcelArray& arr)
      : _base(arr)
      , _refCount(new size_t(1))
    {}

    PyExcelArray::PyExcelArray(const ExcelObj & arr)
      : PyExcelArray(ExcelArray(arr))
    {}

    PyExcelArray::~PyExcelArray()
    {
      *_refCount -= 1;
      if (*_refCount == 0)
      {
        delete _refCount;
        _refCount = nullptr;
      }
    }

    size_t PyExcelArray::refCount() const 
    { 
      return _refCount ? *_refCount : 0;
    }

    const ExcelArray& PyExcelArray::base() const 
    { 
      return _base; 
    }

    py::object PyExcelArray::operator()(size_t row, size_t col) const
    {
      return PySteal<>(
        PyFromAny()(
          _base((ExcelArray::row_t)row, (ExcelArray::col_t)col)));
    }

    py::object PyExcelArray::operator()(size_t row) const
    {
      return PySteal<>(
        PyFromAny()(
          _base((ExcelArray::row_t)row)));
    }

    PyExcelArray PyExcelArray::slice(
      int fromRow, int fromCol, int toRow, int toCol) const
    {
      return PyExcelArray(*this, _base.slice(fromRow, fromCol, toRow, toCol));
    }

    pybind11::object PyExcelArray::getItem(pybind11::tuple loc) const
    {
      if (dims() == 1)
      {
        size_t from, to;
        bool singleElem = getItemIndexReader1d(loc[0], size(), from, to);
        return singleElem
          ? operator()(from)
          : py::cast<PyExcelArray>(slice((int)from, 0, (int)to, 1));
      }
      else
      {
        size_t fromRow, fromCol, toRow, toCol;
        bool singleElem = getItemIndexReader2d(loc, nRows(), nCols(),
          fromRow, fromCol, toRow, toCol);
        return singleElem
          ? operator()(fromRow, fromCol)
          : py::cast<PyExcelArray>(
              PyExcelArray(*this, 
              ExcelArray(_base, 
                (ExcelObj::row_t)fromRow, (ExcelObj::col_t)fromCol,
                (ExcelObj::row_t)toRow,   (ExcelObj::col_t)toCol)));
      }
    }

    size_t PyExcelArray::nRows() const { return _base.nRows(); }
    size_t PyExcelArray::nCols() const { return _base.nCols(); }
    size_t PyExcelArray::size() const { return _base.size(); }
    size_t PyExcelArray::dims() const { return _base.dims(); }

    pybind11::tuple PyExcelArray::shape() const
    {
      if (dims() == 2)
      {
        py::tuple result(2);
        result[0] = nRows();
        result[1] = nCols();
        return result;
      }
      else
      {
        py::tuple result(1);
        result[0] = size();
        return result;
      }
    }

    ExcelType PyExcelArray::dataType() const { return _base.dataType(); }

    auto toArray(const PyExcelArray& arr, std::optional<int> dtype, std::optional<int> dims)
    {
      return PySteal<>(excelArrayToNumpyArray(arr.base(), dims ? *dims : 2, dtype ? *dtype : -1));
    }

    namespace
    {
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind the PyExcelArray type to ExcelArray. PyExcelArray is a wrapper
        // around the core ExcelArray type.
        auto aType = py::class_<PyExcelArray>(mod, "ExcelArray",
          R"(
            A view of a internal Excel array which can be manipulated without
            copying the underlying data. It's not a general purpose array class 
            but rather used to create efficiencies in type converters.
    
            It can be accessed and sliced using the usual syntax (the slice step must be 1):

            ::

                x[1, 1] # The value at 1,1 as int, str, float, etc.

                x[1, :] # The second row as another ExcelArray

                x[:-1, :-1] # A sub-array omitting the last row and column

          )")
          .def("slice", 
            &PyExcelArray::slice, 
            R"(
              Slices the array 
            )",
            py::arg("from_row"), 
            py::arg("from_col"),
            py::arg("to_row"), 
            py::arg("to_col"))
          .def("to_numpy",
            &toArray,
            R"(
              Converts the array to a numpy array. If *dtype* is None, xlOil attempts 
              to determine the correct numpy dtype. It raises an exception if values
              cannot be converted to a specified *dtype*. The array dimension *dims* 
              can be 1 or 2 (default is 2).
            )",
            py::arg("dtype") = py::none(), 
            py::arg("dims") = 2)
          .def("__getitem__", 
            &PyExcelArray::getItem,
            R"(
              Given a 2-tuple, slices the array to return a sub ExcelArray or a single element.
            )")
          .def_property_readonly("nrows", 
            &PyExcelArray::nRows,
            "Returns the number of rows in the array")
          .def_property_readonly("ncols", 
            &PyExcelArray::nCols,
            "Returns the number of columns in the array")
          .def_property_readonly("dims", 
            &PyExcelArray::dims,
            "Property which gives the dimension of the array: 1 or 2")
          .def_property_readonly("shape", 
            &PyExcelArray::shape,
            "Returns a tuple (nrows, ncols) like numpy's array.shape");

        ExcelArrayType = (PyTypeObject*)aType.get_type().ptr();

        mod.def("to_array", &toArray,
          py::arg("array"), 
          py::arg("dtype") = py::none(), 
          py::arg("dims") = 2);
      });
    }
  }
}