#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include "Numpy.h"
#include "BasicTypes.h"
#include "PyCore.h"
#include "ArrayHelpers.h"
#include <xlOil/TypeConverters.h>
#include <xlOil/NumericTypeConverters.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ArrayBuilder.h>
#include <xloil/Date.h>
#include <xloil/StringUtils.h>
#include <numpy/arrayobject.h>
#include <numpy/arrayscalars.h>
#include <numpy/npy_math.h>
#include <numpy/ndarrayobject.h>
#include <pybind11/pybind11.h>
#include <locale>
#include <tuple>

namespace py = pybind11;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::to_string;

// Useful references:
//
// https://docs.scipy.org/doc/numpy/reference/arrays.ndarray.html
//
namespace xloil 
{
  namespace Python
  {
    bool importNumpy()
    {
      auto ret = _import_array();
      return ret == 0;
    }
    bool isArrayDataType(PyTypeObject* t)
    {
      return (t == &PyGenericArrType_Type || PyType_IsSubtype(t, &PyGenericArrType_Type));
    }

    bool isNumpyArray(PyObject * p)
    {
      return PyArray_Check(p);
    }

    /**********************
     * Numpy helper types *
     **********************/

    // We need to override the nan returned here as numpy's nan is not
    // the same as the one defined in numeric_limits for some reason.
    struct ToDoubleNPYNan : conv::ToDouble<double>
    {
      using ToDouble::operator();
      double operator()(CellError err) const
      {
        switch (err)
        {
        case CellError::Div0:
          return NPY_INFINITY;
        case CellError::Num:
        case CellError::Null:
        case CellError::NA:
          return NPY_NAN;
        }
        return ToDouble::operator()(err);
      }
    };

    struct ToFloatNPYNan : public FromExcelBase<float>
    {
      template<class T>
      float operator()(T x) const
      {
        return static_cast<float>(ToDoubleNPYNan()(x));
      }
    };

    class NumpyDateFromDate : public FromExcelBase<npy_datetime>
    {
    public:
      using FromExcelBase::operator();

      npy_datetime operator()(int x) const
      {
        int day, month, year;
        excelSerialDateToYMD(x, year, month, day);
        npy_datetimestruct dt = { year, month, day };
        return PyArray_DatetimeStructToDatetime(NPY_FR_us, &dt);
      }
      npy_datetime operator()(double x) const
      {
        int day, month, year, hours, mins, secs, usecs;
        excelSerialDatetoYMDHMS(x, year, month, day, hours, mins, secs, usecs);
        npy_datetimestruct dt = { year, month, day, hours, mins, secs, usecs };
        return PyArray_DatetimeStructToDatetime(NPY_FR_us, &dt);
      }
      npy_datetime operator()(const PStringView<>& str) const
      {
        std::tm tm;
        if (stringToDateTime(str.view(), tm))
        {
          npy_datetimestruct dt = { tm.tm_year + 1900, tm.tm_mon + 1,
            tm.tm_mday, tm.tm_hour, tm.tm_min, tm.tm_sec, 0 };
          return PyArray_DatetimeStructToDatetime(NPY_FR_us, &dt);
        }
        XLO_THROW("Cannot read '{0}' as a date");
      }
    };
   

    struct TruncateUTF16ToChar
    {
      using to_char = char;
      size_t operator()(to_char* target, size_t size, const wchar_t* begin, const wchar_t* end) const
      {
        auto* p = target;
        auto* pEnd = target + size;
        for (; begin < end && p != pEnd; ++begin, ++p)
          *p = (char)*begin;
        return p - target;
      }
    };

    template <class TConv>
    struct ToFixedWidthString
    {
      using TChar = typename TConv::to_char;
      TConv _conv;
      void operator() (TChar* dest, size_t destSize, const ExcelObj& obj) const
      {
        size_t nWritten = 0;
        auto destLength = destSize / sizeof(TChar);
        if (obj.type() == ExcelType::Str)
        {
          auto pstr = obj.asPString();
          nWritten = _conv(dest, destLength, pstr.begin(), pstr.end());
        }
        else
        {
          auto str = obj.toString();
          nWritten = _conv(dest, destLength, (wchar_t*)str.data(), (wchar_t*)str.data() + str.length());
        }
        memset((char*)(dest + nWritten), 0, (destLength - nWritten) * sizeof(TChar));
      }
    };

    /// <summary>
    /// Helper class which assigns the value of an ExcelObj conversion 
    /// to an numpy array element.  The size_t param is only used for
    /// strings
    /// </summary>
    template<class TExcelObjConverter, class TResultValue>
    struct NPToT
    {
      void operator()(TResultValue* d, size_t, const ExcelObj& x) const
      {
        *d = (TResultValue)FromExcel<TExcelObjConverter>()(x);
      }
    };

    template<class TExcelObjConverter>
    struct NPToT<TExcelObjConverter, PyObject*>
    {
      void operator()(PyObject** d, size_t, const ExcelObj& x) const
      {
        *d = TExcelObjConverter()(x);
      }
    };

    template <class T> struct TypeTraitsBase { };

    template <int> struct TypeTraits {};
    template<> struct TypeTraits<NPY_BOOL> { using storage = bool; using from_excel = NPToT<conv::ToBool<>, storage>;  };
    template<> struct TypeTraits<NPY_SHORT> { using storage = short;  using from_excel = NPToT<conv::ToInt<>, storage>;  };
    template<> struct TypeTraits<NPY_USHORT> { using storage = unsigned short; using from_excel = NPToT<conv::ToInt<>, storage>;  };
    template<> struct TypeTraits<NPY_INT> { using storage = int; using from_excel = NPToT<conv::ToInt<>, storage>; };
    template<> struct TypeTraits<NPY_UINT> { using storage = unsigned; using from_excel = NPToT<conv::ToInt<>, storage>; };
    template<> struct TypeTraits<NPY_LONG> { using storage = long; using from_excel = NPToT<conv::ToInt<>, storage>; };
    template<> struct TypeTraits<NPY_ULONG> { using storage = unsigned long; using from_excel = NPToT<conv::ToInt<>, storage>; };
    template<> struct TypeTraits<NPY_FLOAT> { using storage = float; using from_excel = NPToT<ToFloatNPYNan, storage>; };
    template<> struct TypeTraits<NPY_DOUBLE> { using storage = double; using from_excel = NPToT<ToDoubleNPYNan, storage>; };
    template<> struct TypeTraits<NPY_DATETIME> 
    { 
      using storage = npy_datetime;
      using from_excel = NPToT<NumpyDateFromDate, storage>;
    };
    template<> struct TypeTraits<NPY_STRING> 
    { 
      using from_excel = ToFixedWidthString<TruncateUTF16ToChar>;
      using storage = char;
    };
    template<> struct TypeTraits<NPY_UNICODE> 
    { 
      using from_excel = ToFixedWidthString<ConvertUTF16ToUTF32>;
      using storage = char32_t;
    };
    template<> struct TypeTraits<NPY_OBJECT> 
    { 
      using storage = PyObject * ;
      using from_excel = NPToT<PyFromAny, PyObject*>;
    };

    /// <summary>
    /// Returns the storage size required to write the given array as
    /// a numpy array
    /// </summary>
    template<int TNpType>
    size_t getItemSize(const ExcelArray&)
    {
      return sizeof(TypeTraits<TNpType>::storage);
    }

    template<>
    size_t getItemSize<NPY_STRING>(const ExcelArray& arr)
    {
      size_t strLength = 0;
      for (ExcelArray::size_type i = 0; i < arr.size(); ++i)
        strLength = std::max<size_t>(strLength, arr.at(i).maxStringLength());
      return strLength * sizeof(TypeTraits<NPY_STRING>::storage);
    }

    template<>
    size_t getItemSize<NPY_UNICODE>(const ExcelArray& arr)
    {
      return getItemSize<NPY_STRING>(arr) 
        * sizeof(TypeTraits<NPY_UNICODE>::storage);
    }

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
      case NPY_ULONG:     return TThing<NPY_ULONG>()(std::forward<Args>(args)...);
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

    template <int TNpType>
    class PyFromArray1d : public PyFromExcelImpl
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using data_type = typename TypeTraits<TNpType>::storage;
    public:
      PyFromArray1d(bool trim = true) : _trim(trim) 
      {}

      using PyFromExcelImpl::operator();
      static constexpr char* const ourName = "ndarray(1d)";

      PyObject* operator()(ArrayVal obj) const
      {
        ExcelArray arr(obj.obj, _trim);
        return (*this)(arr);
      }

      PyObject* operator()(const ExcelArray& arr) const
      {
        Py_intptr_t dims[] = { (intptr_t)arr.size() };

        if (arr.size() == 0)
          return PyArray_EMPTY(1, dims, TNpType, 0);

        if (arr.dims() != 1)
          XLO_THROW("Expecting a 1-dim array");

        const auto itemsize = getItemSize<TNpType>(arr);
        const auto dataSize = arr.size() * itemsize;
        auto* data = (char*) PyDataMem_NEW(dataSize);

        auto d = data;
        for (ExcelArray::size_type i = 0; i < arr.size(); ++i, d += itemsize)
          _conv((data_type*)d, itemsize, arr.at(i));
        
        return PyArray_New(
          &PyArray_Type,
          _countof(dims), 
          dims, 
          TNpType, 
          nullptr,  // strides
          data, 
          (int)itemsize, 
          NPY_ARRAY_OWNDATA, 
          nullptr); // array finalizer
      }

      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    template <int TNpType>
    class PyFromArray2d : public PyFromExcelImpl
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using TDataType = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray2d(bool trim = true) : _trim(trim) 
      {}

      using PyFromExcelImpl::operator();
      static constexpr char* const ourName = "ndarray(2d)";

      PyObject* operator()(ArrayVal obj) const
      {
        ExcelArray arr(obj.obj, _trim);
        return (*this)(arr);
      }

      PyObject* operator()(const ExcelArray& arr) const
      {
        // Arrays passed to/from Excel can never be empty but a trimmed 
        // or sliced ExcelArray might be
        if (arr.size() == 0)
        {
          Py_intptr_t dims[] = { 0, 0 };
          return PyArray_EMPTY(2, dims, TNpType, 0);
        }

        Py_intptr_t dims[] = { (intptr_t)arr.nRows(), (intptr_t)arr.nCols() };
        
        const auto itemsize = getItemSize<TNpType>(arr);
        const auto dataSize = arr.size() * itemsize;
        auto data = (char*) PyDataMem_NEW(dataSize);

        auto d = data;
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j, d += itemsize)
            _conv((TDataType*)d, itemsize, arr.at(i, j));

        return PyArray_New(
          &PyArray_Type,
          _countof(dims), 
          dims, 
          TNpType, 
          nullptr, // strides
          data, 
          (int)itemsize,
          NPY_ARRAY_OWNDATA, 
          nullptr); // array inalizer
         
      }

      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    template<
      int TNpType, 
      bool IsString = (TNpType == NPY_UNICODE) || (TNpType == NPY_STRING)>
    struct FromArrayImpl
    {
      using TDataType = typename TypeTraits<TNpType>::storage;

      FromArrayImpl(PyArrayObject* pArr)
      { 
        PyArray_ITEMSIZE(pArr) == sizeof(TDataType) && PyArray_TYPE(pArr) == TNpType;
      }
      static constexpr size_t stringLength = 0;
      auto toExcelObj(
        ExcelArrayBuilder&, 
        void* arrayPtr)
      {
        auto x = (TDataType*)arrayPtr;
        return ExcelObj(*x);
      }
    };

    template <int TNpType>
    struct FromArrayImpl<TNpType, true>
    {
      using data_type = typename TypeTraits<TNpType>::storage;

      // The number of char16 we require to hold any character in the array
      static constexpr uint16_t charMultiple =
        std::max<uint16_t>(1, sizeof(data_type) / sizeof(char16_t));
      
      // Contains the number of characters per numpy array element multiplied 
      // by the number of char16 we will need
      const uint16_t stringLength;

      FromArrayImpl(PyArrayObject* pArr)
        : stringLength(std::min<uint16_t>(USHRT_MAX,
            (uint16_t)PyArray_ITEMSIZE(pArr) / sizeof(data_type) * charMultiple))
      {
        const auto type = PyArray_TYPE(pArr);
        if (type != NPY_UNICODE && type != NPY_STRING)
          XLO_THROW("Incorrect array type: expected string or unicode");
      }

      auto toExcelObj(
        ExcelArrayBuilder& builder, 
        void* arrayPtr)
      {
        const auto x = (const char32_t*)arrayPtr;
        const auto len = strlen32(x, stringLength / charMultiple);
        auto pstr = builder.string((uint16_t)len);
        auto nChars = ConvertUTF32ToUTF16()(
          (char16_t*)pstr.pstr(), pstr.length(), x, x + len );

        // Because not every UTF-32 char takes two UTF-16 chars and not
        // every string takes up the full fixed with, we resize here
        pstr.resize((char16_t)nChars);
        return ExcelObj(std::move(pstr));
      }
    };

    template<>
    struct FromArrayImpl<NPY_OBJECT, false>
    {
      size_t stringLength;

      FromArrayImpl(PyArrayObject* pArr)
      {
        auto type = PyArray_TYPE(pArr);
        if (PyArray_ITEMSIZE(pArr) != sizeof(PyObject*) || type != NPY_OBJECT)
          XLO_THROW("Incorrect array type: expected object");

        stringLength = 0;
        auto dims = PyArray_DIMS(pArr);
        auto nDims = PyArray_NDIM(pArr);

        switch (nDims)
        {
        case 1:
          for (auto i = 0; i < dims[0]; ++i)
            accumulateObjectStringLength(*(PyObject**)PyArray_GETPTR1(pArr, i), stringLength);
          break;
        case 2:
          for (auto i = 0; i < dims[0]; ++i)
            for (auto j = 0; j < dims[1]; ++j)
              accumulateObjectStringLength(*(PyObject**)PyArray_GETPTR2(pArr, i, j), stringLength);
          break;
        default:
          XLO_THROW("FromArray: dimension must be 1 or 2");
        }
      }
      
      static auto toExcelObj(
        ExcelArrayBuilder& builder, 
        void* arrayPtr)
      {
        const auto* pyObj = *(PyObject**)arrayPtr;
        return FromPyObj()(pyObj, builder.charAllocator());
      }
    };

    namespace
    {
      std::tuple<PyArrayObject*, npy_intp*, int, bool> 
        getArrayInfo(const PyObject& obj)
      {
        if (!PyArray_Check(&obj))
          XLO_THROW("Expected array");

        auto pyArr = (PyArrayObject*)&obj;
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
    class XlFromArray1d : public IConvertToExcel<PyObject>
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;

    public:
      XlFromArray1d(bool cache = false) : _cache(cache) {}

      virtual ExcelObj operator()(const PyObject& obj) const override
      {
        auto [pyArr, dims, nDims, isEmpty] = getArrayInfo(obj);
        // Empty arrays are not allowed in Excel, the closest is #N/A.
        if (isEmpty)
          return CellError::NA;

        if (nDims != 1)
          XLO_THROW("Expected 1-d array");
        
        TImpl converter(pyArr);

        ExcelArrayBuilder builder((uint32_t)dims[0], 1, converter.stringLength);
        for (auto j = 0; j < dims[0]; ++j)
          builder(j, 0).emplace(converter.toExcelObj(builder, PyArray_GETPTR1(pyArr, j)));
        
        return _cache
          ? ExcelObj(makeCached<ExcelObj>(builder.toExcelObj()))
          : builder.toExcelObj();
      }
    };

    template <int TNpType>
    class XlFromArray2d : public IConvertToExcel<PyObject>
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;

    public:
      XlFromArray2d(bool cache = false) : _cache(cache) {}

      virtual ExcelObj operator()(const PyObject& obj) const override
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
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
            builder(i, j).emplace(converter.toExcelObj(builder, PyArray_GETPTR2(pyArr, i, j)));

        return _cache
          ? ExcelObj(xloil::makeCached<ExcelObj>(builder.toExcelObj()))
          : builder.toExcelObj();
      }
    };

    int excelTypeToDtype(ExcelType t)
    {
      switch (t)
      {
      case ExcelType::Bool: return NPY_BOOL;
      case ExcelType::Num: return NPY_DOUBLE;
      case ExcelType::Int: return NPY_INT;
      case ExcelType::Str: return NPY_UNICODE;
      default: return NPY_OBJECT;
      }
    }

    PyObject* excelArrayToNumpyArray(const ExcelArray& arr, int dims, int dtype)
    {
      if (dtype < 0)
        dtype = excelTypeToDtype(arr.dataType());
      switch (dims)
      {
      case 1:
        return switchDataType<PyFromArray1d>(dtype, arr);
      case 2:
        return switchDataType<PyFromArray2d>(dtype, arr);
      default:
        XLO_THROW("Dimensions must be 1 or 2");
      };
    }

    ExcelObj numpyArrayToExcel(const PyObject* p)
    {
      if (!PyArray_Check(p))
        XLO_THROW("Expected array");

      auto pArr = (PyArrayObject*)p;
      auto nDims = PyArray_NDIM(pArr);
      auto dType = PyArray_TYPE(pArr);
      switch (nDims)
      {
      case 1:
        return switchDataType<XlFromArray1d>(dType, *p);
      case 2:
        return switchDataType<XlFromArray2d>(dType, *p);
      default:
        XLO_THROW("Expected 1 or 2 dim array");
      }
    }

    namespace
    {
      constexpr const char* name(int numpyDataType)
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

      constexpr const char* dims(int n)
      {
        switch (n)
        {
        case 1: return "_1d";
        case 2: return "_2d";
        default: return "?";
        }
      }

      template <int TNpType>
      using Array1dFromXL = PyFromExcelConverter<PyFromArray1d<TNpType>>;

      template <int TNpType>
      using Array2dFromXL = PyFromExcelConverter<PyFromArray2d<TNpType>>;
  
      template<template<int N> class T, int TNpType, int TNDims>
      struct Reader
      {
        auto operator()(pybind11::module& mod) const
        {
          return py::class_<T<TNpType>, IPyFromExcel, shared_ptr<T<TNpType>>>
            (mod, (prefix + name(TNpType) + dims(TNDims)).c_str())
            .def(py::init<bool>(), py::arg("trim")=true);
        }
        static inline auto prefix = string(theReadConverterPrefix);
      };
      template<template<int N> class T, int TNpType, int TNDims>
      struct Writer
      {
        auto operator()(pybind11::module& mod) const
        {
          return py::class_<T<TNpType>, IPyToExcel, shared_ptr<T<TNpType>>>
            (mod, (prefix + name(TNpType) + dims(TNDims)).c_str())
            .def(py::init<bool>(), py::arg("cache") = false);
        }
        static inline auto prefix = string(theReturnConverterPrefix);
      };
      template<template<template<int> class, int, int> class Func, template<int N> class T, int TNDims>
      void declare(pybind11::module& mod)
      {
        Func<T, NPY_INT,    TNDims>()(mod);
        Func<T, NPY_DOUBLE, TNDims>()(mod);
        Func<T, NPY_BOOL,   TNDims>()(mod);
        Func<T, NPY_STRING, TNDims>()(mod);
        Func<T, NPY_OBJECT, TNDims>()(mod);

        auto datetime = Func<T, NPY_DATETIME, TNDims>()(mod);
        // Alias so that either date or datetime arrays can be requested.
        // TODO: strictly should drop time information if it exists
        mod.add_object((Func<T, 1, 1>::prefix + string("Array_date_") + dims(TNDims)).c_str(), datetime);
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        declare<Reader, Array1dFromXL, 1>(mod);
        declare<Reader, Array2dFromXL, 2>(mod);
        declare<Writer, XlFromArray1d, 1>(mod);
        declare<Writer, XlFromArray2d, 2>(mod);
      });
    }
  }
}