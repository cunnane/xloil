#pragma once


#include "CPython.h"
#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#define PY_ARRAY_UNIQUE_SYMBOL xloil_PyArray_API
#define NO_IMPORT_ARRAY
#include <numpy/arrayobject.h>
#include <numpy/arrayscalars.h>
#include <numpy/npy_math.h>
#undef NO_IMPORT_ARRAY

#include <xlOil/ExcelArray.h>
#include <xlOil/ArrayBuilder.h>
#include <xloil/StringUtils.h>
#include <pybind11/pybind11.h>


namespace xloil
{
  namespace Python
  {
    template <int> struct TypeTraits {};
    template<> struct TypeTraits<NPY_BOOL> { using storage = bool; };
    template<> struct TypeTraits<NPY_SHORT> { using storage = short; };
    template<> struct TypeTraits<NPY_USHORT> { using storage = unsigned short; };
    template<> struct TypeTraits<NPY_INT> { using storage = int; };
    template<> struct TypeTraits<NPY_UINT> { using storage = unsigned; };
    template<> struct TypeTraits<NPY_LONG> { using storage = long; };
    template<> struct TypeTraits<NPY_ULONG> { using storage = unsigned long; };
    template<> struct TypeTraits<NPY_LONGLONG> { using storage = long long; };
    template<> struct TypeTraits<NPY_ULONGLONG> { using storage = unsigned long; };
    template<> struct TypeTraits<NPY_FLOAT> { using storage = float; };
    template<> struct TypeTraits<NPY_DOUBLE> { using storage = double; };
    template<> struct TypeTraits<NPY_DATETIME> { using storage = npy_datetime; };
    template<> struct TypeTraits<NPY_STRING> { using storage = char; };
    template<> struct TypeTraits<NPY_UNICODE> { using storage = char32_t; };
    template<> struct TypeTraits<NPY_OBJECT> { using storage = PyObject*; };

    std::tuple<PyArrayObject*, npy_intp*, int>
      getArrayInfo(const PyObject* obj)
    {
      if (!PyArray_Check(obj))
        XLO_THROW("Expected an array, got a {}", to_string(pybind11::type::of((PyObject*)obj)));

      auto pyArr = (PyArrayObject*)obj;
      auto dims = PyArray_DIMS(pyArr);
      auto nDims = PyArray_NDIM(pyArr);

      return { pyArr, dims, nDims };
    }

    /// C++ safe version of NPY_BEGIN_THREADS_DESCR
    class NumpyBeginThreadsDescr {
    public:
      explicit NumpyBeginThreadsDescr(int dtype)
        : tstate(nullptr)
      {
        if (dtype != NPY_OBJECT)
          tstate = PyEval_SaveThread();
      }

      ~NumpyBeginThreadsDescr()
      {
        if (tstate)
          PyEval_RestoreThread(tstate);
      }

    private:
      PyThreadState* tstate;
    };

    template <template <int> class TThing, class... Args>
    auto switchDataType(int dtype, Args&&... args)
    {
      switch (dtype)
      {
      case NPY_BOOL:      return TThing<NPY_BOOL>()(std::forward<Args>(args)...);
      case NPY_SHORT:     return TThing<NPY_SHORT>()(std::forward<Args>(args)...);
      case NPY_USHORT:    return TThing<NPY_USHORT>()(std::forward<Args>(args)...);
      case NPY_UINT:      return TThing<NPY_UINT>()(std::forward<Args>(args)...);
      case NPY_INT:       return TThing<NPY_INT>()(std::forward<Args>(args)...);
      case NPY_LONG:      return TThing<NPY_LONG>()(std::forward<Args>(args)...);
      case NPY_LONGLONG:  return TThing<NPY_LONGLONG>()(std::forward<Args>(args)...);
      case NPY_ULONG:     return TThing<NPY_ULONG>()(std::forward<Args>(args)...);
      case NPY_ULONGLONG: return TThing<NPY_ULONGLONG>()(std::forward<Args>(args)...);
      case NPY_FLOAT:     return TThing<NPY_FLOAT>()(std::forward<Args>(args)...);
      case NPY_DOUBLE:    return TThing<NPY_DOUBLE>()(std::forward<Args>(args)...);
      case NPY_DATETIME:  return TThing<NPY_DATETIME>()(std::forward<Args>(args)...);
      case NPY_OBJECT:    return TThing<NPY_OBJECT>()(std::forward<Args>(args)...);
      case NPY_STRING:    return TThing<NPY_STRING>()(std::forward<Args>(args)...);
      case NPY_UNICODE:   return TThing<NPY_UNICODE>()(std::forward<Args>(args)...);
      default:
        XLO_THROW("Unsupported numpy date type");
      }
    }

    template<NPY_DATETIMEUNIT TGranularity>
    npy_datetime convertDateTime(const npy_datetimestruct& dt) noexcept;

    template<
      template<template<int> class, int, int> class Declarer,
      template<int N> class Converter,
      int TNDims>
    void declare(pybind11::module& mod)
    {
      Declarer<Converter, NPY_INT, TNDims>()(mod);
      Declarer<Converter, NPY_DOUBLE, TNDims>()(mod);
      Declarer<Converter, NPY_BOOL, TNDims>()(mod);
      Declarer<Converter, NPY_STRING, TNDims>()(mod);
      Declarer<Converter, NPY_OBJECT, TNDims>()(mod);

      auto datetime = Declarer<Converter, NPY_DATETIME, TNDims>()(mod);
      // Alias so that either date or datetime arrays can be requested.
      // TODO: strictly should drop time information if it exists
      mod.add_object(
        (Declarer<Converter, 1, 1>::prefix + string("Array_date_") + dimsToStr(TNDims)).c_str(),
        datetime);
    }

    PyArray_Descr* 
      createDatetimeDtype(int type_num = NPY_DATETIME, NPY_DATETIMEUNIT unit = NPY_FR_us);

    constexpr const char* nameToStr(int numpyDataType)
    {
      switch (numpyDataType)
      {
      case NPY_INT: return "Array_int";
      case NPY_DOUBLE: return "Array_float";
      case NPY_BOOL: return "Array_bool";
      case NPY_DATETIME: return "Array_datetime";
      case NPY_STRING: return "Array_str";
      case NPY_OBJECT: return "Array_object";
      default: return "?";
      }
    }

    constexpr const char* dimsToStr(int n)
    {
      switch (n)
      {
      case 1: return "_1d";
      case 2: return "_2d";
      default: return "_bad_dims_";
      }
    }

    template<int TNpType, typename = void>
    struct FromArrayImpl
    {
      using TDataType = typename TypeTraits<TNpType>::storage;

      FromArrayImpl(PyArrayObject* /*pArr*/)
      {}

      static constexpr size_t stringLength = 0;

      auto toExcelObj(
        ExcelArrayBuilder&,
        void* arrayPtr) const
      {
        auto x = (TDataType*)arrayPtr;
        return ExcelObj(*x);
      }
    };

    template <int TNpType>
    struct FromArrayImpl<TNpType, std::enable_if_t<(TNpType == NPY_FLOAT) || (TNpType == NPY_DOUBLE) || (TNpType == NPY_LONGDOUBLE)>>
    {
      using TDataType = typename TypeTraits<TNpType>::storage;

      FromArrayImpl(PyArrayObject* /*pArr*/)
      {}

      static constexpr size_t stringLength = 0;

      auto toExcelObj(
        ExcelArrayBuilder&,
        void* arrayPtr) const
      {
        auto x = *(TDataType*)arrayPtr;
        if (npy_isnan(x)) return ExcelObj(CellError::NA);
        if (npy_isinf(x)) return ExcelObj(CellError::Num);
        return ExcelObj(x);
      }
    };

    template <int TNpType>
    struct FromArrayImpl<TNpType, std::enable_if_t<(TNpType == NPY_UNICODE) || (TNpType == NPY_STRING)>>
    {
      using data_type = typename TypeTraits<TNpType>::storage;

      // The number of char16 we require to hold any character in the array
      static constexpr uint16_t charMultiple =
        std::max<uint16_t>(1, sizeof(data_type) / sizeof(char16_t));

      // Contains the number of characters per numpy array element multiplied 
      // by the number of char16 we will need
      const size_t stringLength;
      const uint16_t itemLength;

      FromArrayImpl(PyArrayObject* pArr)
        : itemLength(std::min<uint16_t>(
          USHRT_MAX,
          (uint16_t)PyArray_ITEMSIZE(pArr) / sizeof(data_type)))
        , stringLength(charMultiple* itemLength* PyArray_SIZE(pArr))
      {
        const auto type = PyArray_TYPE(pArr);
        if (type != NPY_UNICODE && type != NPY_STRING)
          XLO_THROW("Incorrect array type: expected string or unicode");
      }

      auto toExcelObj(
        ExcelArrayBuilder& builder,
        void* arrayPtr) const
      {
        const auto x = (const char32_t*)arrayPtr;
        const auto len = strlen32(x, itemLength);
        auto pstr = builder.string((uint16_t)len);
        auto nChars = ConvertUTF32ToUTF16()(
          (char16_t*)pstr.pstr(), pstr.length(), x, x + len);

        // Because not every UTF-32 char takes two UTF-16 chars and not
        // every string takes up the full fixed with, we resize here
        pstr.resize((char16_t)nChars);
        return ExcelObj(std::move(pstr));
      }
    };

    template<>
    struct FromArrayImpl<NPY_DATETIME>
    {
      static constexpr size_t stringLength = 0;

      const PyArray_DatetimeMetaData* _meta;

      FromArrayImpl(PyArrayObject* pArr);

      ExcelObj toExcelObj(
        ExcelArrayBuilder& /*builder*/,
        void* arrayPtr) const;
    };
  }
}