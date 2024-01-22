#include "NumpyHelpers.h"
#include "PyCore.h"
#include "BasicTypes.h"
#include <xloil/FPArray.h>

using std::vector;
using std::string;
using std::shared_ptr;
namespace py = pybind11;
using row_t = xloil::ExcelArray::row_t;
using col_t = xloil::ExcelArray::col_t;


namespace xloil
{
  namespace Python
  {
    namespace
    {
      std::tuple<PyArrayObject*, npy_intp*, int, bool>
        getArrayInfo(const PyObject* obj)
      {
        if (!PyArray_Check(obj))
          XLO_THROW("Expected array");

        auto pyArr = (PyArrayObject*)obj;
        auto dims = PyArray_DIMS(pyArr);
        auto nDims = PyArray_NDIM(pyArr);

        // Empty arrays are not allowed in Excel, the closest is #N/A.
        // Regardless of nDims, if any is zero the array is empty.
        bool isEmpty = false;
        for (auto i = 0; i < nDims; ++i)
          if (dims[i] == 0)
            isEmpty = true;

        return { pyArr, dims, nDims, isEmpty };
      }
    }

   
    template <int TNpType>
    class XlFromArray1d : public IPyToExcel
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;
      using TDataType = typename TypeTraits<NPY_OBJECT>::storage;

    public:
      XlFromArray1d(bool cache = false)
        : _cache(cache)
      {}

      ExcelObj operator()(const PyObject* obj) const override
      {
        auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);
        // Empty arrays are not allowed in Excel, the closest is #N/A.
        if (isEmpty)
          return CellError::NA;

        if (nDims != 1)
          XLO_THROW("Expected 1-d array");

        TImpl converter(pyArr);

        ExcelArrayBuilder builder((row_t)dims[0], 1, converter.stringLength);
        auto elementPtr = PyArray_BYTES(pyArr);
        const auto stride = PyArray_STRIDE(pyArr, 0);
        for (auto j = 0; j < dims[0]; ++j, elementPtr += stride)
          builder(j, 0).take(converter.toExcelObj(builder, elementPtr));

        return _cache
          ? makeCached<ExcelObj>(builder.toExcelObj())
          : builder.toExcelObj();
      }
      const char* name() const override
      {
        return "array(1d)";
      }
    };

    template <int TNpType>
    class XlFromArray2d : public IPyToExcel
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;

    public:
      XlFromArray2d(bool cache = false) : _cache(cache) {}

      ExcelObj operator()(const PyObject* obj) const override
      {
        auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);
        // Empty arrays are not allowed in Excel, the closest is #N/A.
        if (isEmpty)
          return CellError::NA;

        if (nDims != 2)
          XLO_THROW("Expected 2-d array");

        TImpl converter(pyArr);

        ExcelArrayBuilder builder((uint32_t)dims[0], (uint32_t)dims[1],
          converter.stringLength);

        const auto stride1 = PyArray_STRIDE(pyArr, 0);
        const auto stride2 = PyArray_STRIDE(pyArr, 1);
        for (auto i = 0; i < dims[0]; ++i)
        {
          auto elementPtr = PyArray_BYTES(pyArr) + i * stride1;
          for (auto j = 0; j < dims[1]; ++j, elementPtr += stride2)
            builder(i, j).take(converter.toExcelObj(builder, elementPtr));
        }
        return _cache
          ? xloil::makeCached<ExcelObj>(builder.toExcelObj())
          : builder.toExcelObj();
      }
      const char* name() const override
      {
        return "array(2d)";
      }
    };

    // TODO: support converting unknown objects to string. It's easy but
    // how best to pass the parameter?
    template <>
    class XlFromArray1d<NPY_OBJECT> : public IPyToExcel
    {
      bool _cacheResult;
      using TDataType = typename TypeTraits<NPY_OBJECT>::storage;

    public:
      XlFromArray1d(bool cache = false) : _cacheResult(cache) {}

      ExcelObj operator()(const PyObject* obj) const override
      {
        auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);
        // Empty arrays are not allowed in Excel, the closest is #N/A.
        if (isEmpty)
          return CellError::NA;

        if (nDims != 1)
          XLO_THROW("Expected 1-d array");

        size_t stringLength = dims[0] * 4; // err why?

        SequentialArrayBuilder builder((row_t)dims[0], 1, stringLength);
        auto elementPtr = PyArray_BYTES(pyArr);
        const auto stride = PyArray_STRIDE(pyArr, 0);
        for (auto i = 0; i < dims[0]; ++i, elementPtr += stride)
          builder.emplace(
            FromPyObj<detail::ReturnToCache, true>()(
              *(TDataType*)elementPtr, builder.charAllocator()));

        py::gil_scoped_release noGil;

        return _cacheResult
          ? xloil::makeCached<ExcelObj>(builder.toExcelObj())
          : builder.toExcelObj();
      }
      const char* name() const override
      {
        return "array(1d)";
      }
    };

    template <>
    class XlFromArray2d<NPY_OBJECT> : public IPyToExcel
    {
      bool _cacheResult;
      using TDataType = typename TypeTraits<NPY_OBJECT>::storage;

    public:
      XlFromArray2d(bool cache = false) : _cacheResult(cache) {}

      ExcelObj operator()(const PyObject* obj) const override
      {
        auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);
        // Empty arrays are not allowed in Excel, the closest is #N/A.
        if (isEmpty)
          return CellError::NA;

        if (nDims != 2)
          XLO_THROW("Expected 2-d array");

        size_t stringLength = dims[0] * dims[1] * 4; // err why?

        SequentialArrayBuilder builder((row_t)dims[0], (col_t)dims[1], stringLength);
        auto charAllocator = builder.charAllocator();

        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
          {
            auto p = *(TDataType*)PyArray_GETPTR2(pyArr, i, j);
            builder.emplace(FromPyObj<detail::ReturnToCache, true>()(p, charAllocator));
          }

        py::gil_scoped_release noGil;

        return _cacheResult
          ? xloil::makeCached<ExcelObj>(builder.toExcelObj())
          : builder.toExcelObj();
      }
      const char* name() const override
      {
        return "array(2d)";
      }
    };

    ExcelObj numpyArrayToExcel(const PyObject* p)
    {
      if (!PyArray_Check(p))
        XLO_THROW("Expected array");

      auto pArr = (PyArrayObject*)p;
      auto nDims = PyArray_NDIM(pArr);
      auto dType = PyArray_TYPE(pArr);

      NumpyBeginThreadsDescr releaseGil(dType);

      switch (nDims)
      {
      case 1:
        return switchDataType<XlFromArray1d>(dType, p);
      case 2:
        return switchDataType<XlFromArray2d>(dType, p);
      default:
        XLO_THROW("Expected 1 or 2 dim array");
      }
    }
  
    std::shared_ptr<FPArray> numpyToFPArray(const PyObject* obj)
    {
      auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);

      if (isEmpty || nDims != 2)
        XLO_THROW("Expected non-empty 2-d array");

      if (PyArray_TYPE(pyArr) != NPY_DOUBLE)
        XLO_THROW("Expected float array (type float64)");

      const auto itemsize = PyArray_ITEMSIZE(pyArr);
      const auto strides = PyArray_STRIDES(pyArr);

      auto result = FPArray::create(dims[0], dims[1]);

      // Check if the array is in row-major order like the FPArray so we can
      // use memcpy (note the strides are in bytes).
      if (strides[0] == itemsize * dims[0] && strides[1] == itemsize)
      {
        const auto* raw = PyArray_BYTES(pyArr);
        const auto databytes = itemsize * dims[0] * dims[1];
        memcpy(result->begin(), raw, databytes);
      }
      else
      {
        // Do it the old-fashioned way: elementwise copy
        auto pArr = result->begin();
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
            *pArr++ = *(double*)PyArray_GETPTR2(pyArr, i, j);
      }

      return result;
    }
    namespace
    {
      template<template<int N> class T, int TNpType, int TNDims>
      struct Writer
      {
        auto operator()(pybind11::module& mod) const
        {
          return py::class_<T<TNpType>, IPyToExcel, shared_ptr<T<TNpType>>>
            (mod, (prefix + nameToStr(TNpType) + dimsToStr(TNDims)).c_str())
            .def(py::init<bool>(), py::arg("cache") = false);
        }
        static inline auto prefix = string(theReturnConverterPrefix);
      };

      static int theBinder = addBinder([](py::module& mod)
      {
        declare<Writer, XlFromArray1d, 1>(mod);
        declare<Writer, XlFromArray2d, 2>(mod);
      });
    }
  }
}