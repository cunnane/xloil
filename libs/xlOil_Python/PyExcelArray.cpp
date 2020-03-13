#pragma once
#include "PyHelpers.h"
#include "PyExcelArray.h"
#include "BasicTypes.h"

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    PyExcelArray::PyExcelArray(const PyExcelArray& from, int fromRow, int fromCol, int toRow, int toCol)
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

    py::object PyExcelArray::operator()(int row, int col) const
    {
      return PySteal<>(PyFromExcel<PyFromAny<>>()(_base(row, col)));
    }
    py::object PyExcelArray::operator()(int row) const
    {
      return PySteal<>(PyFromExcel<PyFromAny<>>()(_base(row)));
    }

    PyExcelArray PyExcelArray::subArray(int fromRow, int fromCol, int toRow, int toCol) const
    {
      return PyExcelArray(*this, fromRow, fromCol, toRow, toCol);
    }

    pybind11::object PyExcelArray::getItem(pybind11::tuple loc) const
    {
      if (loc.size() != 2)
        XLO_THROW("Expecting tuple of size 2");
      auto r = loc[0];
      auto c = loc[1];
      size_t fromRow, fromCol, toRow, toCol;
      if (r.is_none())
      {
        fromRow = 0;
        toRow = nRows();
      }
      else if (PySlice_Check(r.ptr()))
      {
        size_t step, sliceLength;
        r.cast<py::slice>().compute(nRows(), &fromRow, &toRow, &step, &sliceLength);
      }
      else
      {
        fromRow = r.cast<size_t>();
        toRow = fromRow + 1;
      }

      if (r.is_none())
      {
        fromCol = 0;
        toCol = nRows();
      }
      else if (PySlice_Check(c.ptr()))
      {
        size_t step, sliceLength;
        c.cast<py::slice>().compute(nCols(), &fromCol, &toCol, &step, &sliceLength);
      }
      else
      {
        fromCol = c.cast<size_t>();
        // Check for single element access
        if (fromRow == toRow + 1)
          return operator()(fromRow, fromCol);
        toCol = fromCol + 1;
      }

      return py::cast<PyExcelArray>(subArray(fromRow, fromCol, toRow, toCol));
    }
    size_t PyExcelArray::nRows() const { return _base.nRows(); }
    size_t PyExcelArray::nCols() const { return _base.nCols(); }
    size_t PyExcelArray::size() const { return _base.size(); }
    size_t PyExcelArray::dims() const { return _base.dims(); }

    ExcelType PyExcelArray::dataType() const { return _base.dataType(); }
  }
}