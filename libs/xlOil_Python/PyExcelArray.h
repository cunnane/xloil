#pragma once
#include <xlOil/ExcelArray.h>

namespace xloil
{
  namespace Python
  {
    extern PyTypeObject* ExcelArrayType;

    class PyExcelArray
    {
      ExcelArray _base;
      size_t* _refCount;

      PyExcelArray(
        const PyExcelArray& from,
        ExcelArray&& rebase);

    public:
      PyExcelArray(const PyExcelArray& from);
      PyExcelArray(ExcelArray&& arr);
      PyExcelArray(const ExcelArray& arr);
      ~PyExcelArray();

      /// <summary>
      /// This ref count is a safety feature to stop python code from
      /// keeping references to PyExcelArrays after the function which
      /// created them has exited and hence the underlying array data
      /// is destroyed
      /// </summary>
      /// <returns></returns>
      size_t refCount() const;
      const ExcelArray& base() const;

      pybind11::object operator()(size_t row, size_t col) const;
      pybind11::object operator()(size_t row) const;

      PyExcelArray subArray(int fromRow, int fromCol, int toRow, int toCol) const;
      pybind11::object getItem(pybind11::tuple) const;
      size_t nRows() const;
      size_t nCols() const;
      size_t size() const;
      size_t dims() const;

      /// <summary>
      /// Consistent with numpy's shape property, but read-only
      /// </summary>
      pybind11::tuple shape() const;

      ExcelType dataType() const;
    };

    auto toArray(const PyExcelArray& arr, std::optional<int> dtype, std::optional<int> dims);
  }
}
