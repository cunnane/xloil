#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include "Numpy.h"
#include "BasicTypes.h"
#include "TypeConverters.h"
#include "StandardConverters.h"
#include "ExcelArray.h"
#include "ArrayHelpers.h"
#include "ArrayBuilder.h"
#include <xloil/Date.h>
#include <xloil/StringUtils.h>
#include <numpy/arrayobject.h>
#include <numpy/arrayscalars.h>
#include <numpy/npy_math.h>
#include <numpy/ndarrayobject.h>
#include <pybind11/pybind11.h>
#include <locale>

namespace py = pybind11;
using std::shared_ptr;
using std::unique_ptr;

namespace xloil 
{
  namespace Python
  {
    bool importNumpy()
    {
      import_array();
      return true;
    }
    bool isArrayDataType(PyTypeObject* t)
    {
      return (t == &PyGenericArrType_Type || PyType_IsSubtype(t, &PyGenericArrType_Type));
    }

    bool isNumpyArray(PyObject * p)
    {
      return PyArray_Check(p);
    }

    PyObject* excelToNumpyArray(const ExcelObj& obj, PyTypeObject* type)
    {
      auto descr = PyArray_DescrFromTypeObject((PyObject*)type);
      PyTypeNum_ISINTEGER(descr->type_num);
      return nullptr;
    }


    /**********************
     * Numpy helper types *
     **********************/

    // We need to override the nan returned here as numpy's nan is not
    // the same as defined in numeric_limits for some reason.
    struct ToDoubleNPYNan : ToDouble
    {
      double fromError(CellError err) const
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
        return ToDouble::fromError(err);
      }
    };

    class NumpyDateFromDate : public FromExcelBase<npy_datetime>
    {
    public:
      npy_datetime fromInt(int x) const
      {
        int day, month, year;
        excelSerialDateToDMY(x, day, month, year);
        npy_datetimestruct dt = { year, month, day };
        return PyArray_DatetimeStructToDatetime(NPY_FR_us, &dt);
      }
      npy_datetime fromDouble(double x) const
      {
        int day, month, year, hours, mins, secs, usecs;
        excelSerialDatetoDMYHMS(x, day, month, year, hours, mins, secs, usecs);
        npy_datetimestruct dt = { year, month, day, hours, mins, secs, usecs };
        return PyArray_DatetimeStructToDatetime(NPY_FR_us, &dt);
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
          auto pstr = obj.asPascalStr();
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


    template<class T, class R>
    struct NPToT
    {
      void operator()(R* d, size_t, const ExcelObj& x) const
      {
        *d = T()(x);
      }
    };

    template <class T> struct TypeTraitsBase { };

    template <int> struct TypeTraits {};
    template<> struct TypeTraits<NPY_BOOL> { using storage = bool; using from_excel = NPToT<ToBool, storage>;  };
    template<> struct TypeTraits<NPY_SHORT> { using storage = short;  using from_excel = NPToT<ToInt, storage>;  };
    template<> struct TypeTraits<NPY_USHORT> { using storage = unsigned short; using from_excel = NPToT<ToInt, storage>;  };
    template<> struct TypeTraits<NPY_INT> { using storage = int; using from_excel = NPToT<ToInt, storage>; };
    template<> struct TypeTraits<NPY_UINT> { using storage = unsigned; using from_excel = NPToT<ToInt, storage>; };
    template<> struct TypeTraits<NPY_LONG> { using storage = long; using from_excel = NPToT<ToInt, storage>; };
    template<> struct TypeTraits<NPY_ULONG> { using storage = unsigned long; using from_excel = NPToT<ToInt, storage>; };
    template<> struct TypeTraits<NPY_FLOAT> { using storage = float; using from_excel = NPToT<ToDoubleNPYNan, storage>; };
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
      using from_excel = NPToT<PyFromExcel<PyFromAny<>>, storage>;
    };

    template<int TNpType>
    size_t getItemSize(const ExcelArray& arr)
    {
      return sizeof(TypeTraits<TNpType>::storage);
    }

    template<>
    size_t getItemSize<NPY_STRING>(const ExcelArray& arr)
    {
      size_t strLength = 0;
      for (auto i = 0; i < arr.size(); ++i)
        strLength = std::max(strLength, arr.at(i).maxStringLength());
      return strLength * sizeof(TypeTraits<NPY_STRING>::storage);
    }

    template<>
    size_t getItemSize<NPY_UNICODE>(const ExcelArray& arr)
    {
      return getItemSize<NPY_STRING>(arr) * sizeof(TypeTraits<NPY_UNICODE>::storage);
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
    class PyFromArray1d : public PyFromCache<PyFromArray1d<TNpType>>
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using data_type = typename TypeTraits<TNpType>::storage;
    public:
      PyFromArray1d(bool trim) : _trim(trim) {}

      PyObject* fromArray(const ExcelObj& obj) const
      {
        ExcelArray arr(obj, _trim);
        return fromArrayObj(arr);
      }
      PyObject* fromArrayObj(const ExcelArray& arr) const
      {
        if (arr.size() == 0)
          Py_RETURN_NONE;

        if (arr.dims() != 1)
          XLO_THROW("Expecting a 1-dim array");

        Py_intptr_t dims[1];
        dims[0] = arr.size();
        const int nDims = 1;

        const auto itemsize = getItemSize<TNpType>(arr);
        const auto dataSize = arr.size() * itemsize;
        auto* data = (char*) PyDataMem_NEW(dataSize);

        auto d = data;
        for (auto i = 0; i < arr.size(); ++i, d += itemsize)
          _conv((data_type*)d, itemsize, arr.at(i));
        
        return PyArray_New(
          &PyArray_Type,
           nDims, dims, TNpType, NULL, data, (int)itemsize, NPY_ARRAY_OWNDATA, NULL);
      }
    };

    template <int TNpType>
    class PyFromArray2d : public PyFromCache<PyFromArray2d<TNpType>>
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using TDataType = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray2d(bool trim) : _trim(trim) {}

      PyObject* fromArray(const ExcelObj& obj) const
      {
        ExcelArray arr(obj, _trim);
        return fromArrayObj(arr);
      }

      PyObject* fromArrayObj(const ExcelArray& arr) const
      {
        if (arr.size() == 0)
          Py_RETURN_NONE;

        Py_intptr_t dims[2];
        const int nDims = 2;
        dims[0] = arr.nRows();
        dims[1] = arr.nCols();

        const auto itemsize = getItemSize<TNpType>(arr);
        const auto dataSize = arr.size() * itemsize;
        auto data = (char*) PyDataMem_NEW(dataSize);

        auto d = data;
        for (auto i = 0; i < arr.nRows(); ++i)
          for (auto j = 0; j < arr.nCols(); ++j, d += itemsize)
            _conv((TDataType*)d, itemsize, arr.at(i, j));

        return PyArray_New(
          &PyArray_Type,
          nDims, dims, TNpType, NULL, data, (int)itemsize, NPY_ARRAY_OWNDATA, NULL);
      }
    };

    template <int TNpType, bool IsString = (TNpType == NPY_UNICODE) || (TNpType == NPY_STRING)>
    struct FromArrayImpl
    {
      using TDataType = typename TypeTraits<TNpType>::storage;

      FromArrayImpl(PyArrayObject* pArr)
      { 
        PyArray_ITEMSIZE(pArr) == sizeof(TDataType) && PyArray_TYPE(pArr) == TNpType;
      }
      static constexpr size_t stringLength = 0;
      void builderEmplace(ExcelArrayBuilder& b, size_t i, size_t j, void* arrayPtr)
      {
        auto x = (TDataType*)arrayPtr;
        b.emplace_at(i, j, *x);
      }
    };

    template <int TNpType>
    struct FromArrayImpl<TNpType, true>
    {
      using data_type = typename TypeTraits<TNpType>::storage;
      static constexpr size_t charMultiple = TNpType == NPY_UNICODE ? 2 : 1;
      FromArrayImpl(PyArrayObject* pArr)
      {
        const auto type = PyArray_TYPE(pArr);
        if (type != NPY_UNICODE && type != NPY_STRING)
          XLO_THROW("Incorrect array type");
        stringLength = std::min<size_t>(USHRT_MAX, PyArray_ITEMSIZE(pArr) / sizeof(data_type) * charMultiple);
      }
      size_t stringLength;

      void builderEmplace(ExcelArrayBuilder& builder, size_t i, size_t j, void* arrayPtr)
      {
        auto x = (char32_t*)arrayPtr;
        PString<> pstr((char16_t)stringLength);
        auto nChars = ConvertUTF32ToUTF16()(
          (char16_t*)pstr.pstr(), pstr.length(), x, x + stringLength / charMultiple);
        pstr.resize((char16_t)nChars);
        builder.emplace_at(i, j, pstr);
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
          XLO_THROW("Incorrect array type");

        stringLength = 0;
        auto p = *(PyObject**)pArr;
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
        default:
          XLO_THROW("FromArray: dimension must be 1 or 2");
        }
      }
      
      static void builderEmplace(ExcelArrayBuilder& builder, size_t i, size_t j, void* arrayPtr)
      {
        auto* x = *(PyObject**)arrayPtr;
        FromPyObj()(x, [&builder, i, j](auto&&... args) { return builder.emplace_at(i, j, args...); });
      }
    };

    template <int TNpType>
    class XlFromArray1d : public IConvertToExcel<PyObject>
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;

    public:
      XlFromArray1d(bool cache = false) : _cache(cache) {}

      virtual ExcelObj operator()(const PyObject& obj) const override
      {
        if (!PyArray_Check(&obj))
          XLO_THROW("Expected array");

        auto pyArr = (PyArrayObject*)&obj;
        auto dims = PyArray_DIMS(pyArr);
        auto nDims = PyArray_NDIM(pyArr);
        if (nDims != 1)
          XLO_THROW("Expected 1-d array");
        
        TImpl converter(pyArr);

        ExcelArrayBuilder builder(1, dims[0], converter.stringLength);
        for (auto j = 0; j < dims[0]; ++j)
          converter.builderEmplace(builder, 0, j, PyArray_GETPTR1(pyArr, j));
        
        return _cache
          ? theCore->insertCache(builder.toExcelObj())
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
        if (!PyArray_Check(&obj))
          XLO_THROW("Expected array");

        auto pyArr = (PyArrayObject*)&obj;
        auto dims = PyArray_DIMS(pyArr);
        auto nDims = PyArray_NDIM(pyArr);
        if (nDims != 2)
          XLO_THROW("Expected 2-d array");

        TImpl converter(pyArr);

        ExcelArrayBuilder builder(dims[0], dims[1], converter.stringLength);
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
            converter.builderEmplace(builder, i, j, PyArray_GETPTR2(pyArr, i, j));

        return _cache
          ? theCore->insertCache(builder.toExcelObj())
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
    template<int Dtype, template<int> class TWhat>
    struct ThingyImpl
    {
      auto operator()(const ExcelArray& arr) const { return TWhat<Dtype>(true).fromArrayObj(arr); }
    };
    template<int Dtype> using Thingy1 = ThingyImpl<Dtype, PyFromArray1d>;
    template<int Dtype> using Thingy2 = ThingyImpl<Dtype, PyFromArray2d>;

    PyObject* excelArrayToNumpyArray(const ExcelArray& arr, int dims, int dtype)
    {
      if (dtype < 0)
        dtype = excelTypeToDtype(arr.dataType());
      switch (dims)
      {
      case 1:
        return switchDataType<Thingy1>(dtype, arr);
      case 2:
        return switchDataType<Thingy2>(dtype, arr);
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
      template <int TNpType>
      using Array1dFromXL = PyFromExcel<PyFromArray1d<TNpType>>;

      template <int TNpType>
      using Array2dFromXL = PyFromExcel<PyFromArray2d<TNpType>>;

      template<class T>
      void declare(pybind11::module& mod, const char* type)
      {
        py::class_<T, IPyFromExcel, shared_ptr<T>>
          (mod, ("To_Array_" + std::string(type)).c_str())
          .def(py::init<bool>(), py::arg("trim")=true);
      }

      //template<class T>
      //void declare2(pybind11::module& mod, const char* name)
      //{
      //  py::class_<T, IPyToExcel, shared_ptr<T>>(mod, name)
      //    .def(py::init<bool>(), py::arg("cache") = false);
      //}
      static int theBinder = addBinder([](py::module& mod)
      {
        declare<Array1dFromXL<NPY_INT>      >(mod, "int_1d");
        declare<Array1dFromXL<NPY_DOUBLE>   >(mod, "float_1d");
        declare<Array1dFromXL<NPY_BOOL>     >(mod, "bool_1d");
        declare<Array1dFromXL<NPY_DATETIME> >(mod, "date_1d");
        declare<Array1dFromXL<NPY_STRING>   >(mod, "str_1d");
        declare<Array1dFromXL<NPY_OBJECT>   >(mod, "object_1d");

        declare<Array2dFromXL<NPY_INT>      >(mod, "int_2d");
        declare<Array2dFromXL<NPY_DOUBLE>   >(mod, "float_2d");
        declare<Array2dFromXL<NPY_BOOL>     >(mod, "bool_2d");
        declare<Array2dFromXL<NPY_DATETIME> >(mod, "date_2d");
        declare<Array2dFromXL<NPY_STRING>   >(mod, "str_2d");
        declare<Array2dFromXL<NPY_OBJECT>   >(mod, "object_2d");

        //declare2<XlFromArray1d<EmplacePyObj<FromPyObj>, PyObject*, NPY_OBJECT > > (mod, "Array_object_1d_to_Excel");
        //declare2<XlFromArray1d<EmplaceType<double>, double, NPY_DOUBLE > >(mod, "Array_float_1d_to_Excel");
      });
    }
  }
}