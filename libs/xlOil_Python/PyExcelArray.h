#pragma once
#include "ExcelArray.h"

namespace xloil
{
  namespace Python
  {
    class PyExcelArray
    {
      ExcelArray _base;
      size_t* _refCount;

      PyExcelArray(const PyExcelArray& from, int fromRow, int fromCol, int toRow, int toCol);

    public:
      PyExcelArray(const PyExcelArray& from);
      PyExcelArray(ExcelArray&& arr);
      PyExcelArray(const ExcelArray& arr);
      ~PyExcelArray();

      size_t refCount() const;
      const ExcelArray& base() const;

      pybind11::object operator()(int row, int col) const;
      pybind11::object operator()(int row) const;

      PyExcelArray subArray(int fromRow, int fromCol, int toRow, int toCol) const;
      pybind11::object getItem(pybind11::tuple) const;
      size_t nRows() const;
      size_t nCols() const;
      size_t size() const;
      size_t dims() const;

      ExcelType dataType() const;
    };
  }
}
