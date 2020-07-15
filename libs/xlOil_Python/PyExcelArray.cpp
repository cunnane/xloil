#pragma once
#include "PyHelpers.h"
#include "PyExcelArray.h"
#include "BasicTypes.h"

using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    extern PyTypeObject* ExcelArrayType = nullptr;

    PyExcelArray::PyExcelArray(
      const PyExcelArray& from, 
      size_t fromRow, size_t fromCol,
      int toRow, int toCol)
      : _base(ExcelArray(from._base, fromRow, fromCol, toRow, toCol))
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

    PyExcelArray::~PyExcelArray()
    {
      *_refCount -= 1;
      if (_refCount == 0)
        delete _refCount;
    }
    size_t PyExcelArray::refCount() const { return *_refCount; }
    const ExcelArray& PyExcelArray::base() const { return _base; }

    py::object PyExcelArray::operator()(size_t row, size_t col) const
    {
      return PySteal<>(PyFromExcel<PyFromAny<>>()(_base(row, col)));
    }
    py::object PyExcelArray::operator()(size_t row) const
    {
      return PySteal<>(PyFromExcel<PyFromAny<>>()(_base(row)));
    }

    PyExcelArray PyExcelArray::subArray(size_t fromRow, size_t fromCol, int toRow, int toCol) const
    {
      return PyExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    pybind11::object PyExcelArray::getItem(pybind11::tuple loc) const
    {
      size_t fromRow, fromCol, toRow, toCol;
      bool single = sliceHelper(loc, nRows(), nCols(),
        fromRow, fromCol, toRow, toCol);
      return single
        ? operator()(fromRow, fromCol)
        : py::cast<PyExcelArray>(subArray(fromRow, fromCol, (int)toRow, (int)toCol));
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
        auto aType = py::class_<PyExcelArray>(mod, "ExcelArray")
          .def("sub_array", &PyExcelArray::subArray, 
            py::arg("from_row"), 
            py::arg("from_col"),
            py::arg("to_row") = 0, 
            py::arg("to_col") = 0)
          .def("to_numpy", &toArray,
            py::arg("dtype") = py::none(), 
            py::arg("dims") = 2)
          .def("__getitem__", &PyExcelArray::getItem)
          .def_property_readonly("nrows", &PyExcelArray::nRows)
          .def_property_readonly("ncols", &PyExcelArray::nCols)
          .def_property_readonly("dims", &PyExcelArray::dims)
          .def_property_readonly("shape", &PyExcelArray::shape);

        ExcelArrayType = (PyTypeObject*)aType.get_type().ptr();

        mod.def("to_array", &toArray,
          py::arg("array"), 
          py::arg("dtype") = py::none(), 
          py::arg("dims") = 2);
      }, 100);
    }
  }
}