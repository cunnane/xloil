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
#include <xloil/FPArray.h>
#include <xloil/StringUtils.h>
#include <numpy/arrayobject.h>
#include <numpy/arrayscalars.h>
#include <numpy/npy_math.h>
#include <numpy/ndarrayobject.h>
#include <pybind11/pybind11.h>
#include <locale>
#include <tuple>

using std::vector;

typedef npy_int64 npy_datetime;

// Prior to the as-yet-unrealased Numpy 2, there are no working API functions
// to convert to and from numpy's datetime representation. There are
// promising looking API functions, but they give a ten-year-old deprecation
// error. The only approach is to copy/paste the relevant conversion code
// which is what we have done here
namespace
{
#include "numpy_datetime.c"
}

namespace
{
  template<NPY_DATETIMEUNIT TGranularity>
  auto convertDateTime(const npy_datetimestruct& dt) noexcept
  {
    PyArray_DatetimeMetaData meta{ TGranularity , 1 };
    npy_datetime result;
    NpyDatetime_ConvertDatetimeStructToDatetime64(&meta, &dt, &result);
    return result;
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
}

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
    struct ToDoubleNPYNan : conv::ExcelValToType<double, double>
    {
      using base = conv::ExcelValToType<double, double>;
      using base::operator();
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
        return base::operator()(err);
      }
    };

    struct ToFloatNPYNan : public ExcelValVisitor<float>
    {
      template<class T>
      float operator()(T x) const
      {
        return static_cast<float>(ToDoubleNPYNan()(x));
      }
    };

    class NumpyDateFromDate : public ExcelValVisitor<npy_datetime>
    {
    public:
      using ExcelValVisitor::operator();

      npy_datetime operator()(int x) const noexcept
      {
        int day, month, year;
        excelSerialDateToYMD(x, year, month, day);
        npy_datetimestruct dt{ year, month, day, 0, 0, 0, 0, 0 };
        return convertDateTime<NPY_FR_us>(dt);
      }
      npy_datetime operator()(double x) const noexcept
      {
        int day, month, year, hours, mins, secs, usecs;
        excelSerialDatetoYMDHMS(x, year, month, day, hours, mins, secs, usecs);
        npy_datetimestruct dt{ year, month, day, hours, mins, secs, usecs };
        return convertDateTime<NPY_FR_us>(dt);
      }
      npy_datetime operator()(const PStringRef& str) const
      {
        std::tm tm;
        if (stringToDateTime(str.view(), tm))
        {
          npy_datetimestruct dt{ tm.tm_year + 1900, tm.tm_mon + 1,
            tm.tm_mday, tm.tm_hour, tm.tm_min, tm.tm_sec, 0 };
          return convertDateTime<NPY_FR_us>(dt);
        }
        XLO_THROW("Cannot read '{0}' as a date");
      }
    };
   
    double excelDateFromNumpyDate(const npy_datetime x, const PyArray_DatetimeMetaData& meta)
    {
      npy_datetimestruct dt;
      NpyDatetime_ConvertDatetime64ToDatetimeStruct(
        const_cast<PyArray_DatetimeMetaData*>(&meta), x, &dt);

      if (dt.year == NPY_DATETIME_NAT)
        return NPY_NAN;

      if (meta.base <= NPY_FR_D)
        return excelSerialDateFromYMD(dt.year, dt.month, dt.day);
      else
        return excelSerialDateFromYMDHMS(
          dt.year, dt.month, dt.day, dt.hour, dt.min, dt.sec, dt.us);
    }

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
          auto pstr = obj.cast<PStringRef>();
          nWritten = _conv(dest, destLength, pstr.begin(), pstr.end());
        }
        else
        {
          auto str = obj.toStringRecursive();
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
        *d = (TResultValue)x.visit(TExcelObjConverter());
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
    template<> struct TypeTraits<NPY_BOOL>   { using storage = bool;           using from_excel = NPToT<conv::ToType<bool>, storage>; };
    template<> struct TypeTraits<NPY_SHORT>  { using storage = short;          using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_USHORT> { using storage = unsigned short; using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_INT>    { using storage = int;            using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_UINT>   { using storage = unsigned;       using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_LONG>   { using storage = long;           using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_ULONG>  { using storage = unsigned long;  using from_excel = NPToT<conv::ToType<int>, storage>; };
    template<> struct TypeTraits<NPY_FLOAT>  { using storage = float;          using from_excel = NPToT<ToFloatNPYNan, storage>; };
    template<> struct TypeTraits<NPY_DOUBLE> { using storage = double;         using from_excel = NPToT<ToDoubleNPYNan, storage>; };
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
    

    /// <summary>
    /// Helper to call PyArray_New.  Allocate data using PyDataMem_NEW
    /// </summary>
    template<int NDim>
    PyObject* newNumpyArray(int numpyType, Py_intptr_t (&dims)[NDim], void* data, size_t itemsize)
    {
      if (numpyType == NPY_DATETIME)
      {
        PyArray_DatetimeMetaData meta{ NPY_FR_us, 1 };
        auto descr = create_datetime_dtype(NPY_DATETIME, &meta);
        return PyArray_NewFromDescr(
          &PyArray_Type,
          descr,
          NDim,
          dims,
          nullptr, // strides
          data,
          NPY_ARRAY_OWNDATA,
          nullptr); // array finaliser
      }
      else
      {
        return PyArray_New(
          &PyArray_Type,
          NDim,
          dims,
          numpyType,
          nullptr, // strides
          data,
          (int)itemsize,
          NPY_ARRAY_OWNDATA,
          nullptr); // array finaliser
      }
    }

    template <int TNpType>
    class PyFromArray1d : public detail::PyFromExcelImpl
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using data_type = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray1d(bool trim = true) : _trim(trim), _conv()
      {}

      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "array(1d)";

      PyObject* operator()(const ArrayVal& obj) const
      {
        ExcelArray arr(obj, _trim);
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
        const auto arraySize = arr.size();

        // If array memory size is over 65k and the string length is
        // 64 chars, just switch to using an array of object strings
        if (itemsize > 256 && arraySize * itemsize > 1 << 16)
        {
          return PyFromArray1d<NPY_OBJECT>(_trim)(arr);
        }
        else
        {
          auto data = (char*)PyDataMem_NEW(arraySize * itemsize);
          auto d = data;
          for (auto p = arr.begin(); p != arr.end(); ++p, d += itemsize)
            _conv((data_type*)d, itemsize, *p);
          return newNumpyArray(TNpType, dims, data, itemsize);
        }
      }

      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    template <int TNpType>
    class PyFromArray2d : public detail::PyFromExcelImpl
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using TDataType = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray2d(bool trim = true) : _trim(trim), _conv()
      {}

      using detail::PyFromExcelImpl::operator();
      static constexpr char* const ourName = "array(2d)";

      PyObject* operator()(const ArrayVal& obj) const
      {
        ExcelArray arr(obj, _trim);
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
        const auto arraySize = arr.size();

        // If array memory size is over 65k and the string length is
        // 64 chars, just switch to using an array of object strings
        if (itemsize > 256 && arraySize * itemsize > 1 << 16)
        {
          return PyFromArray1d<NPY_OBJECT>(_trim)(arr);
        }
        else
        {
          const auto dataSize = arraySize * itemsize;
          auto data = (char*)PyDataMem_NEW(dataSize);

          auto d = data;
          for (auto i = 0; i < dims[0]; ++i)
          {
            auto* pObj = arr.row_begin(i);
            const auto* rowEnd = pObj + dims[1];
            for (; pObj != rowEnd; d += itemsize, ++pObj)
              _conv((TDataType*)d, itemsize, *pObj);
          }

          return newNumpyArray(TNpType, dims, data, itemsize);
        }
      }

      constexpr wchar_t* failMessage() const { return L"Expected array"; }
    };

    PyObject* numpyArrayFromCArray(size_t rows, size_t columns, const double* array)
    {
      Py_intptr_t dims[] = { (intptr_t)rows, (intptr_t)columns };

      constexpr auto itemsize = sizeof(double);
      const auto dataSize = rows * columns * itemsize;

      auto data = (char*)PyDataMem_NEW(dataSize);

      memcpy(data, array, dataSize);

      return newNumpyArray(NPY_DOUBLE, dims, data, itemsize);
    }

    class FPArrayConverter : public IPyFromExcel
    {
    public:
      virtual PyObject* operator()(
        const ExcelObj& xl, const PyObject* /*defaultVal*/) override
      {
        auto& fp = reinterpret_cast<const msxll::FP12&>(xl);
        return numpyArrayFromCArray(fp.rows, fp.columns, fp.array);
      }
      const char* name() const override
      {
        return "FloatArray";
      }
    };

    IPyFromExcel* createFPArrayConverter()
    {
      return new FPArrayConverter();
    }

    template<
      int TNpType, 
      bool IsString = (TNpType == NPY_UNICODE) || (TNpType == NPY_STRING),
      bool IsFloat = (TNpType == NPY_FLOAT) || (TNpType == NPY_DOUBLE) || (TNpType == NPY_LONGDOUBLE)>
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
    struct FromArrayImpl<TNpType, false, true>
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
    struct FromArrayImpl<TNpType, true, false>
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
        void* arrayPtr) const
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
      
      auto toExcelObj(
        ExcelArrayBuilder& builder, 
        void* arrayPtr) const
      {
        const auto* pyObj = *(PyObject**)arrayPtr;
        return FromPyObj()(pyObj, builder.charAllocator());
      }
    };

    template<>
    struct FromArrayImpl<NPY_DATETIME, false>
    {
      static constexpr size_t stringLength = 0;

      const PyArray_DatetimeMetaData* _meta;

      FromArrayImpl(PyArrayObject* pArr)
        : _meta(get_datetime_metadata_from_dtype(PyArray_DESCR(pArr)))
      {}

      auto toExcelObj(
        ExcelArrayBuilder& /*builder*/,
        void* arrayPtr) const
      {
        auto x = (npy_datetime*)arrayPtr;
        const auto serial = excelDateFromNumpyDate(*x, *_meta);
        return ExcelObj(serial);
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
    class XlFromArray1d : public IPyToExcel
    {
      bool _cache;
      using TImpl = FromArrayImpl<TNpType>;

    public:
      XlFromArray1d(bool cache = false) 
        : _cache(cache) 
      {}

      ExcelObj operator()(const PyObject& obj) const override
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

      ExcelObj operator()(const PyObject& obj) const override
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
          ? xloil::makeCached<ExcelObj>(builder.toExcelObj())
          : builder.toExcelObj();
      }
      const char* name() const override
      {
        return "array(2d)";
      }
    };

    shared_ptr<FPArray> numpyToFPArray(const PyObject& obj)
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

    int excelTypeToNumpyDtype(ExcelType t)
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
        dtype = excelTypeToNumpyDtype(arr.dataType());

      NumpyBeginThreadsDescr releaseGil(dtype);

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

      NumpyBeginThreadsDescr releaseGil(dType);

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

    
    PyObject* toNumpyDatetimeFromExcelDateArray(const PyObject* obj)
    {
      if (!PyArray_Check(obj))
        XLO_THROW("Expected array");

      const auto array = (PyArrayObject*)obj;

      npy_uint32 op_flags[2];
      /*
       * No inner iteration - inner loop is handled by CopyArray code
       */
      auto flags = NPY_ITER_EXTERNAL_LOOP; // use NPY_ITER_C_INDEX rather?
      /*
       * Tell the constructor to automatically allocate the output.
       * The data type of the output will match that of the input.
       */
      PyArrayObject* op[] = { array, nullptr };
      op_flags[0] = NPY_ITER_READONLY;
      op_flags[1] = NPY_ITER_WRITEONLY | NPY_ITER_ALLOCATE;

      PyArray_DatetimeMetaData meta{ NPY_FR_us, 1 };
      auto outputDescr = PySteal((PyObject*)create_datetime_dtype(NPY_DATETIME, &meta));
      PyArray_Descr* op_descr[] = { PyArray_DescrFromType(NPY_DOUBLE), (PyArray_Descr*)outputDescr.ptr()};

      /* Construct the iterator */
      auto iter = NpyIter_MultiNew(2, op, flags, NPY_KEEPORDER, NPY_SAFE_CASTING,
        op_flags, op_descr);
      if (!iter)
        XLO_THROW("Failed to create iterator: expected numeric array");

      {
        NumpyBeginThreadsDescr dropGil(NPY_DOUBLE);

        /*
         * Make a copy of the iternext function pointer and
         * a few other variables the inner loop needs.
         */
        auto iternext = NpyIter_GetIterNext(iter, NULL);
        auto innerstride = NpyIter_GetInnerStrideArray(iter)[0];
        auto itemsize = NpyIter_GetDescrArray(iter)[0]->elsize;
        /*
         * The inner loop size and data pointers may change during the
         * loop, so just cache the addresses.
         */
        auto innersizeptr = NpyIter_GetInnerLoopSizePtr(iter);
        auto dataptrarray = NpyIter_GetDataPtrArray(iter);

        /*
         * Note that because the iterator allocated the output,
         * it matches the iteration order and is packed tightly,
         * so we don't need to check it like the input.
         */

         /* For efficiency, should specialize this based on item size... */
        do {
          npy_intp N = *innersizeptr;
          const char* in = dataptrarray[0];
          char* out = dataptrarray[1];

          for (npy_intp i = 0; i < N; i++)
          {
            *((npy_datetime*)out) = NumpyDateFromDate()(*(double*)in);
            in += innerstride;
            out += itemsize;
          }

        } while (iternext(iter));
      }

      /* Get the result from the iterator object array */
      auto ret = NpyIter_GetOperandArray(iter)[1];
      Py_INCREF(ret);

      if (NpyIter_Deallocate(iter) != NPY_SUCCEED) 
      {
        Py_DECREF(ret);
        XLO_THROW("Failed to deallocate iterator");
      }

      return (PyObject*)ret;
    }

    namespace TableHelpers
    {
      struct ApplyConverter
      {
        virtual ~ApplyConverter() {}
        virtual void operator()(ExcelArrayBuilder& builder,
          xloil::detail::ArrayBuilderIterator& start,
          xloil::detail::ArrayBuilderIterator& end) const = 0;
      };

      template<int NPDtype>
      struct ConverterHolder : public ApplyConverter
      {
        FromArrayImpl<NPDtype> _impl;
        PyArrayObject* _array;

        ConverterHolder(PyArrayObject* array)
          : _impl(array)
          , _array(array)
        {}

        virtual ~ConverterHolder() {}

        size_t stringLength() const { return _impl.stringLength; }

        virtual void operator()(ExcelArrayBuilder& builder,
          xloil::detail::ArrayBuilderIterator& start,
          xloil::detail::ArrayBuilderIterator& end) const
        {
          char* arrayPtr = PyArray_BYTES(_array);
          const auto step = PyArray_STRIDES(_array)[0];
          for (; start != end; arrayPtr += step, ++start)
          {
            start->emplace(_impl.toExcelObj(builder, arrayPtr));
          }
        }
      };

      /// <summary>
      /// Helper class used with `switchDataType`
      /// </summary>
      template<int NPDtype>
      struct CreateConverter
      {
        ApplyConverter* operator()(PyArrayObject* array, size_t& stringLength)
        {
          auto converter = new ConverterHolder<NPDtype>(array);
          stringLength += converter->stringLength();
          return converter;
        }
      };

      size_t arrayShape(const py::object& p)
      {
        if (p.is_none())
          return 0;

        if (!PyArray_Check(p.ptr()))
          XLO_THROW("Expected an array");

        auto pyArr = (PyArrayObject*)p.ptr();
        const auto nDims = PyArray_NDIM(pyArr);
        const auto dims = PyArray_DIMS(pyArr);

        if (nDims != 1)
          XLO_THROW("Expected 1 dim array");

        return dims[0];
      }

      /// <summary>
      /// This class holds an array of virtual FromArrayImpl holders.  Each column in a 
      /// dataframe can have a different data type and so require a different converter.
      /// The indices can also have their own data types. The class uses `collect` to 
      /// examine 1-d numpy arrays and creates an appropriate converters. Then `write` is
      /// called when and ExcelArrayBuilder object is ready to receive the converted data
      /// </summary>
      struct Converters
      {
        vector<unique_ptr<const ApplyConverter>> _converters;
        size_t stringLength = 0;
        bool _hasObjectDtype;

        Converters(size_t n)
        {
          _converters.reserve(n);
          _hasObjectDtype = false;
        }

        auto collect(const py::object& p, size_t expectedLength)
        {
          auto shape = arrayShape(p);
          
          if (shape != expectedLength)
            XLO_THROW("Expected a 1-dim array of size {}", expectedLength);

          auto pyArr = (PyArrayObject*)p.ptr();
          const auto dtype = PyArray_TYPE(pyArr);
          if (dtype == NPY_OBJECT)
            _hasObjectDtype = true;

          _converters.emplace_back(unique_ptr<ApplyConverter>(
            switchDataType<CreateConverter>(dtype, pyArr, std::ref(stringLength))));
        }

        auto write(size_t iArray, ExcelArrayBuilder& builder, int startX, int startY, bool byRow)
        {
          auto start = byRow
            ? builder.row_begin(startX) + startY
            : builder.col_begin(startX) + startY;

          auto end = byRow
            ? builder.row_end(startX)
            : builder.col_end(startX);

          (*_converters[iArray])(builder, start, end);
        }

        /// <summary>
        /// Used to determine if we can release the GIL for the duration of the conversion
        /// </summary>
        auto hasObjectDtype() const { return _hasObjectDtype; }
      };
    }

    ExcelObj numpyTableHelper(
      uint32_t nOuter,
      uint32_t nInner,
      const py::object& columns,
      const py::object& rows,
      const py::object& headings,
      const py::object& index,
      const py::object& indexName)
    {
      // This method can handle writing data vertically or horizontally.  When used to 
      // write a pandas DataFrame, the data is vertical/by-column.
      const auto byRow = columns.is_none();

      const auto hasHeadings = !headings.is_none();
      const auto hasIndex = !index.is_none();

      auto tableData = const_cast<PyObject*>(byRow ? rows.ptr() : columns.ptr());

      // The row or column headings can be multi-level indices. We determine the number
      // of levels from iterators later.
      auto nHeadings = 0;
      auto nIndex = 0;

      py::object iter;
      PyObject* item;

      // Converters may end up larger if we have multi-level indices
      TableHelpers::Converters converters(
        nOuter + (hasHeadings ? 1 : 0) + (hasIndex ? 1 : 0));
      
      // Examine data frame index
      if (hasIndex > 0)
      {
        iter = PySteal(PyObject_GetIter(index.ptr()));
        while ((item = PyIter_Next(iter.ptr())) != 0)
        {
          converters.collect(PySteal(item), nInner);
          ++nIndex;
        }
      }

      iter = PySteal(PyObject_GetIter(tableData));
      // First loop to establish array size and length of strings
      while ((item = PyIter_Next(iter.ptr())) != 0)
      {
        converters.collect(PySteal(item), nInner);
      }

      if (hasHeadings > 0)
      {
        iter = PySteal(PyObject_GetIter(headings.ptr()));
        while ((item = PyIter_Next(iter.ptr())) != 0)
        {
          converters.collect(PySteal(item), nOuter);
          ++nHeadings;
        }
      }

      vector<ExcelObj> indexNames(nIndex * nHeadings, CellError::NA);
      auto indexNameStringLength = 0;
      if (nIndex > 0 && !indexName.is_none())
      {
        iter = PySteal(PyObject_GetIter(indexName.ptr()));
        auto i = 0;
        while (i < nIndex * nHeadings && (item = PyIter_Next(iter.ptr())) != 0)
        {
          indexNames[i] = FromPyObj()(PySteal(item).ptr());
          indexNameStringLength += indexNames[i].stringLength();
          ++i;
        }
      }

      iter = py::object(); // Ensure iterator is closed

      // If possible, release the GIL before beginning the conversion
      NumpyBeginThreadsDescr releaseGil(
        converters.hasObjectDtype() ? NPY_OBJECT : NPY_FLOAT);

      auto nRows = nOuter + nIndex;
      auto nCols = nInner + nHeadings;
      if (!byRow)
        std::swap(nRows, nCols);


      ExcelArrayBuilder builder(
        nRows,
        nCols,
        converters.stringLength + indexNameStringLength);

      // Write the index names in the top left
      if (!byRow)
      {
        for (auto j = 0u; j < nIndex; ++j)
          for (auto i = 0u; i < nHeadings; ++i)
            builder(i, j) = indexNames[i * nHeadings + j];
      }
      else
      {
        for (auto j = 0u; j < nIndex; ++j)
          for (auto i = 0u; i < nHeadings; ++i)
            builder(j, i) = indexNames[i * nHeadings + j];
      }

      
      auto iConv = 0;

      for (auto i = 0u; i < nOuter + nIndex; ++i, ++iConv)
        converters.write(iConv, builder, i, nHeadings, byRow);

      for (auto i = 0u; i < nHeadings; ++i, ++iConv)
        converters.write(iConv, builder, i, nIndex, !byRow);

      return builder.toExcelObj();
    }

    namespace
    {
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
            (mod, (prefix + nameToStr(TNpType) + dimsToStr(TNDims)).c_str())
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
            (mod, (prefix + nameToStr(TNpType) + dimsToStr(TNDims)).c_str())
            .def(py::init<bool>(), py::arg("cache") = false);
        }
        static inline auto prefix = string(theReturnConverterPrefix);
      };

      template<
        template<template<int> class, int, int> class Declarer, 
        template<int N> class Converter, 
        int TNDims>
      void declare(pybind11::module& mod)
      {
        Declarer<Converter, NPY_INT,    TNDims>()(mod);
        Declarer<Converter, NPY_DOUBLE, TNDims>()(mod);
        Declarer<Converter, NPY_BOOL,   TNDims>()(mod);
        Declarer<Converter, NPY_STRING, TNDims>()(mod);
        Declarer<Converter, NPY_OBJECT, TNDims>()(mod);

        auto datetime = Declarer<Converter, NPY_DATETIME, TNDims>()(mod);
        // Alias so that either date or datetime arrays can be requested.
        // TODO: strictly should drop time information if it exists
        mod.add_object(
          (Declarer<Converter, 1, 1>::prefix + string("Array_date_") + dimsToStr(TNDims)).c_str(),
          datetime);
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        declare<Reader, Array1dFromXL, 1>(mod);
        declare<Reader, Array2dFromXL, 2>(mod);
        declare<Writer, XlFromArray1d, 1>(mod);
        declare<Writer, XlFromArray2d, 2>(mod);

        mod.def("_table_converter",
          &numpyTableHelper,
          R"(
            For internal use. Converts a table like object (such as a pandas DataFrame) to 
            RawExcelValue suitable for returning to xlOil.
            
            n, m:
              the number of data fields and the length of the fields
            columns / rows: 
              a iterable of numpy array containing data, specified as columns 
              or rows (not both)
            headings:
              optional array of data field headings
            index:
              optional data field labels - one per data point
            index_name:
              optional headings for the index, should be a 1 dim iteratable of size
              num_index_levels * num_column_levels
          )",
          py::arg("n"),
          py::arg("m"),
          py::arg("columns") = py::none(),
          py::arg("rows") = py::none(),
          py::arg("headings") = py::none(),
          py::arg("index") = py::none(),
          py::arg("index_name") = py::none());
      });
    }
  }
}
