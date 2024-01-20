#include "NumpyHelpers.h"
#include "PyCore.h"
#include "BasicTypes.h"
#include <xloil/Date.h>

using std::vector;
namespace py = pybind11;
using std::shared_ptr;
using std::unique_ptr;
using std::string;
using std::to_string;
using row_t = xloil::ExcelArray::row_t;
using col_t = xloil::ExcelArray::col_t;

typedef npy_int64 npy_datetime;


// Useful references:
//
// https://docs.scipy.org/doc/numpy/reference/arrays.ndarray.html
//
namespace xloil 
{
  namespace Python
  {
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
      //float operator()(CellError err) const
      //{

      //}
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

    template <int, typename=void> struct FromExcel {};
    template <int Type> struct FromExcel<Type, std::enable_if_t<std::is_integral_v<typename TypeTraits<Type>::storage>>>
    {
      using value = NPToT<conv::ToType<int>, typename TypeTraits<Type>::storage>;
    };
    template<> struct FromExcel<NPY_BOOL>      { using value = NPToT<conv::ToType<bool>, TypeTraits<NPY_BOOL>::storage>; };
    template<> struct FromExcel<NPY_FLOAT>     { using value = NPToT<ToFloatNPYNan, TypeTraits<NPY_FLOAT>::storage>; };
    template<> struct FromExcel<NPY_DOUBLE>    { using value = NPToT<ToDoubleNPYNan, TypeTraits<NPY_DOUBLE>::storage>; };
    template<> struct FromExcel<NPY_DATETIME>
    {
      using value= NPToT<NumpyDateFromDate, TypeTraits<NPY_DATETIME>::storage>;
    };
    template<> struct FromExcel<NPY_STRING>
    {
      using value= ToFixedWidthString<TruncateUTF16ToChar>;
    };
    template<> struct FromExcel<NPY_UNICODE>
    {
      using value= ToFixedWidthString<ConvertUTF16ToUTF32>;
    };
    template<> struct FromExcel<NPY_OBJECT>
    {
      using value= NPToT<PyFromAny, PyObject*>;
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
      // Start with a length of 1 since this is the minimum itemsize,
      // even for an array of empty strings
      size_t strLength = 1;
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

   

    /// <summary>
    /// Helper to call PyArray_New.  Allocate data using PyDataMem_NEW
    /// </summary>
    template<int NDim>
    PyObject* newNumpyArray(int numpyType, Py_intptr_t (&dims)[NDim], void* data, size_t itemsize)
    {
      if (numpyType == NPY_DATETIME)
      {
        auto descr = createDatetimeDtype();
        return PyArray_NewFromDescr(
          &PyArray_Type,
          descr,
          NDim,
          dims,
          nullptr, // strides
          data,
          NPY_ARRAY_OWNDATA | NPY_ARRAY_CARRAY,
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
          NPY_ARRAY_OWNDATA | NPY_ARRAY_CARRAY,
          nullptr); // array finaliser
      }
    }

    template <int TNpType>
    class PyFromArray1d : public detail::PyFromExcelImpl<PyFromArray1d<TNpType>>
    {
      bool _trim;
      typename FromExcel<TNpType>::value _conv;
      using data_type = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray1d(bool trim = true) : _trim(trim), _conv()
      {}

      using detail::PyFromExcelImpl<PyFromArray1d<TNpType>>::operator();
      static constexpr char* const ourName = "array(1d)";

      PyObject* operator()(const ExcelObj& obj) const
      {
        ExcelArray arr(cacheCheck(obj), _trim);
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
    class PyFromArray2d : public detail::PyFromExcelImpl<PyFromArray2d<TNpType>>
    {
      bool _trim;
      typename FromExcel<TNpType>::value _conv;
      using TDataType = typename TypeTraits<TNpType>::storage;

    public:
      PyFromArray2d(bool trim = true) : _trim(trim), _conv()
      {}

      using detail::PyFromExcelImpl<PyFromArray2d<TNpType>>::operator();
      static constexpr char* const ourName = "array(2d)";

      PyObject* operator()(const ExcelObj& obj) const
      {
        ExcelArray arr(cacheCheck(obj), _trim);
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


      auto outputDescr = PySteal((PyObject*)createDatetimeDtype());
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

    namespace
    {
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
      });
    }
  }
}
