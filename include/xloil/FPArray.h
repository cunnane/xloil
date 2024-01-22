#pragma once
#include <xlOil/XlCallSlim.h>
#include <xlOil/ExportMacro.h>
#include <cassert>
#include <memory>

namespace xloil
{
  /// <summary>
  /// Wraps a FP12 struct: this is a 2-d array of doubles which is an optional
  /// argument type for user-defined functions. It is very fast and lightweight
  /// but is less flexible and user-friendly: if any value in the array passed
  /// to the function is not a number, Excel will return VALUE! without invoking
  /// the function.  A function returning an FPArray, cannot return error conditions
  /// so these must be logged.  See for example Excel's MINVERSE function.
  /// </summary>
  class FPArray : public msxll::_FP12
  {
  private:
    FPArray() {}

  public:
    /// <summary>
    /// Creates a managed `FPArray` that is deleted when AfterCalculate is invoked. You
    /// should return a raw `FPArray*` from your function and not hold a reference.
    /// Since an FPArray is variable size struct, it cannot be created using a normal
    /// C++ constructor. The `FPArray` must be managed as there is no xlAutoFree
    /// callback for Excel `FP12` return types.
    /// </summary>
    XLOIL_EXPORT static std::shared_ptr<FPArray> create(size_t nRows, size_t nCols);

    /// <summary>
    /// Returns the 'empty' static array (actually a 1x1 array containing a NaN) as
    /// a truly empty FPArray is not valid.
    /// </summary>
    /// <returns></returns>
    XLOIL_EXPORT static FPArray* empty();

    /// <summary>
    /// Assigns a the given double to all elements of the array
    /// </summary>
    /// <param name="val"></param>
    /// <returns></returns>
    FPArray& operator=(double val)
    {
      for (auto i = begin(); i != end(); ++i)
        *i = val;
    }
    bool operator==(const FPArray& that)
    {
      return rows == that.rows
        && columns == that.columns
        && std::equal(begin(), end(), that.begin());
    }
    /// <summary>
    /// Total number of elements: rows x columns
    /// </summary>
    /// <returns></returns>
    size_t size() const
    {
      return rows * columns;
    }
    /// <summary>
    /// Retrieves the i-th element (data is stored in column-major order)
    /// </summary>
    double& operator[](size_t i)
    {
      assert(i < size());
      return array[i];
    }
    /// <summary>
    /// Retrieves the i-th element (data is stored in column-major order)
    /// </summary>
    double operator[](size_t i) const
    {
      assert(i < size());
      return array[i];
    }
    /// <summary>
    /// Retrieves the (i,j)-th element
    /// </summary>
    double& operator()(size_t i, size_t j)
    {
      assert(i < (size_t)rows && j < (size_t)columns);
      return operator[](i * columns + j);
    }
    /// <summary>
    /// Retrieves the (i,j)-th element
    /// </summary>
    double operator()(size_t i, size_t j) const
    {
      assert(i < (size_t)rows && j < (size_t)columns);
      return operator[](i * columns + j);
    }
    double* begin()
    {
      return array;
    }
    double* end()
    {
      return array + size();
    }
    const double* begin() const
    {
      return array;
    }
    const double* end() const
    {
      return array + size();
    }
  };
}