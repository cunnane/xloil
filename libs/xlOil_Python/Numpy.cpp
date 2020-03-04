#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include "Numpy.h"
#include "BasicTypes.h"
#include "TypeConverters.h"
#include "StandardConverters.h"
#include "ExcelArray.h"
#include "ArrayHelpers.h"
#include "ArrayBuilder.h"
#include <xloil/Date.h>
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

    template <int>
    struct TypeTraits {};
    template<> struct TypeTraits<NPY_BOOL> { using from_excel = ToBool; using storage_type = bool; };
    template<> struct TypeTraits<NPY_SHORT> { using from_excel = ToInt; using storage_type = short; };
    template<> struct TypeTraits<NPY_USHORT> { using from_excel = ToInt; using storage_type = unsigned short; };
    template<> struct TypeTraits<NPY_INT> { using from_excel = ToInt; using storage_type = int; };
    template<> struct TypeTraits<NPY_UINT> { using from_excel = ToInt; using storage_type = unsigned; };
    template<> struct TypeTraits<NPY_LONG> { using from_excel = ToInt; using storage_type = long; };
    template<> struct TypeTraits<NPY_ULONG> { using from_excel = ToInt; using storage_type = unsigned long; };
    template<> struct TypeTraits<NPY_FLOAT> { using from_excel = ToDoubleNPYNan; using storage_type = float; };
    template<> struct TypeTraits<NPY_DOUBLE> { using from_excel = ToDoubleNPYNan; using storage_type = double; };
    template<> struct TypeTraits<NPY_DATETIME> 
    { 
      using from_excel = NumpyDateFromDate;
      using storage_type = npy_datetime; 
    };
    template<> struct TypeTraits<NPY_STRING> 
    { 
      using from_excel = PyFromExcel<PyFromString>; 
      using to_excel = FromPyString;
      using storage_type = PyObject*; 
    };
    template<> struct TypeTraits<NPY_UNICODE> 
    { 
      using from_excel = PyFromExcel<PyFromString>;
      using to_excel = FromPyString;
      using storage_type = PyObject*; 
    };
    template<> struct TypeTraits<NPY_OBJECT> 
    { 
      using from_excel = PyFromExcel<PyFromAny<>>;
      using to_excel = FromPyObj;
      using storage_type = PyObject*; 
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

    void excelObjToFixedWidthString(char* dest, size_t destSize, const ExcelObj& obj)
    {
      const char* from = 0;
      memset(dest, 0, destSize);
      if (obj.type() == ExcelType::Str)
      {
        auto pstr = obj.asPascalStr();
        auto from = pstr.pstr();
        auto to = from + pstr.length();
        for (; from != to; ++from, dest += 4)
          *(wchar_t*)dest = *from;
      }
      else
      {
        auto str = obj.toString();
        auto from = str.data();
        auto to = from + str.length();
        for (; from != to; ++from, dest += 4)
          *(wchar_t*)dest = *from;
      }
    }

   /* std::wstring_convert<
      std::codecvt_utf16<char32_t, 0x10ffff, std::little_endian>,
      char32_t> theUtf16ToUnicode;

 */
    template <int TNpType>
    class PyFromArray1d : public PyFromCache<PyFromArray1d<TNpType>>
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using TDataType = typename TypeTraits<TNpType>::storage_type;

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
        auto pyObj = PySteal<py::object>(PyArray_SimpleNewFromDescr(nDims, dims, PyArray_DescrFromType(TNpType)));
        auto pyArr = (PyArrayObject*)pyObj.ptr();
        for (auto i = 0; i < arr.size(); ++i)
        {
          TDataType val = _conv(arr.at(i));
          auto ptr = (TDataType*)PyArray_GETPTR1(pyArr, i);
          *ptr = val;
        }
        return pyObj.release().ptr();
      }
    };

    template <>
    class PyFromArray1d<NPY_UNICODE> : public PyFromCache<PyFromArray1d<NPY_UNICODE>>
    {
      bool _trim;
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

        size_t strLength = 0;
        for (auto i = 0; i < arr.size(); ++i)
          strLength = std::max(strLength, arr.at(i).maxStringLength());

        /* NumPy Unicode is always 4-byte */
        auto itemsize = strLength * 4;
        auto dataSize = arr.size() * itemsize;
        void* data = PyDataMem_NEW(dataSize);

        auto d = (char*)data;
        for (auto i = 0; i < arr.size(); ++i, d += itemsize)
          excelObjToFixedWidthString(d, itemsize, arr.at(i));
        
        return PyArray_New(
          &PyArray_Type,
           nDims, dims, NPY_UNICODE, NULL, data, itemsize, NPY_ARRAY_OWNDATA, NULL);
      }
    };

    template <int TNpType>
    class PyFromArray2d : public PyFromCache<PyFromArray2d<TNpType>>
    {
      bool _trim;
      typename TypeTraits<TNpType>::from_excel _conv;
      using TDataType = typename TypeTraits<TNpType>::storage_type;

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


        auto itemsize = sizeof(TDataType);
        auto dataSize = arr.size() * itemsize;
        void* data = PyDataMem_NEW(dataSize);

        auto d = (TDataType*)data;
        for (auto i = 0; i < arr.nRows(); ++i)
          for (auto j = 0; j < arr.nCols(); ++j)
            *d++ = _conv(arr.at(i, j));

        return PyArray_New(
          &PyArray_Type,
          nDims, dims, TNpType, NULL, data, itemsize, NPY_ARRAY_OWNDATA, NULL);

        //auto pyObj = PySteal<py::object>(PyArray_SimpleNewFromDescr(nDims, dims, PyArray_DescrFromType(TNpType)));
        //auto pyArr = (PyArrayObject*)pyObj.ptr();
        //for (auto i = 0; i < arr.nRows(); ++i)
        //  for (auto j = 0; j < arr.nCols(); ++j)
        //  {
        //    TDataType val = _conv(arr.at(i, j));
        //    auto ptr = (TDataType*)PyArray_GETPTR2(pyArr, i, j);
        //    *ptr = val;
        //  }
        //return pyObj.release().ptr();
      }
    };

    template <>
    class PyFromArray2d<NPY_UNICODE> : public PyFromCache<PyFromArray2d<NPY_UNICODE>>
    {
      bool _trim;
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

        if (arr.dims() != 2)
          XLO_THROW("Expecting a 2-dim array");

        Py_intptr_t dims[2];
        const int nDims = 2;
        dims[0] = arr.nRows();
        dims[1] = arr.nCols();

        size_t strLength = 0;
        for (auto i = 0; i < arr.nRows(); ++i)
          for (auto j = 0; j < arr.nCols(); ++j)
            strLength = std::max(strLength, arr.at(i, j).maxStringLength());

        /* NumPy Unicode is always 4-byte */
        auto itemsize = strLength * 4;
        auto dataSize = arr.size() * itemsize;
        void* data = PyDataMem_NEW(dataSize);

        auto d = (char*)data;
          
        for (auto i = 0; i < arr.nRows(); ++i)
          for (auto j = 0; j < arr.nCols(); ++j, d += itemsize)
            excelObjToFixedWidthString(d, itemsize, arr.at(i, j));
          
        return PyArray_New(
          &PyArray_Type,
          nDims, dims, NPY_UNICODE, NULL, data, 4, NPY_ARRAY_OWNDATA, NULL);
      }
    };

    template <int TNpType, bool IsObject = 
      (TNpType==NPY_OBJECT || TNpType == NPY_STRING || TNpType == NPY_UNICODE)>
    struct FromArrayImpl
    {
      using TDataType = typename TypeTraits<TNpType>::storage_type;

      static bool checkType(PyArrayObject* pArr)
      { 
        return PyArray_ITEMSIZE(pArr) == sizeof(TDataType) && PyArray_TYPE(pArr) == TNpType;
      }
      static void addStringLength(void* arrayPtr, size_t& strLength)  {}
      static void builderEmplace(ExcelArrayBuilder& b, size_t i, size_t j, void* arrayPtr)
      {
        auto x = (TDataType*)arrayPtr;
        b.emplace_at(i, j, *x);
      }
    };

    template<int TNpType>
    struct FromArrayImpl<TNpType, true>
    {
      using TConv = typename TypeTraits<TNpType>::to_excel;

      static bool checkType(PyArrayObject* pArr)
      {
        auto type = PyArray_TYPE(pArr);
        return PyArray_ITEMSIZE(pArr) == sizeof(PyObject*) &&
          (type == NPY_OBJECT || type == NPY_UNICODE);
      }
      static void addStringLength(void* arrayPtr, size_t& strLength)
      {
        auto p = *(PyObject**)arrayPtr;
        accumulateObjectStringLength(p, strLength);
      }
      static void builderEmplace(ExcelArrayBuilder& builder, size_t i, size_t j, void* arrayPtr)
      {
        auto x = *(PyObject**)arrayPtr;
        TConv()(x, [&builder, i, j](auto&&... args) { return builder.emplace_at(i, j, args...); });
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
        
        if (!TImpl::checkType(pyArr))
          XLO_THROW("Array data type does not match converter type");

        size_t stringLength = 0;
        for (auto j = 0; j < dims[0]; ++j)
          TImpl::addStringLength(PyArray_GETPTR1(pyArr, j), stringLength);

        ExcelArrayBuilder builder(1, dims[0], stringLength);
        for (auto j = 0; j < dims[0]; ++j)
          TImpl::builderEmplace(builder, 0, j, PyArray_GETPTR1(pyArr, j));
        
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

        if (!TImpl::checkType(pyArr))
          XLO_THROW("Array data type does not match converter type");
        
        // TODO: can the compiler optimise this to nothing when length() returns zero always?
        size_t stringLength = 0;
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
            TImpl::addStringLength(PyArray_GETPTR2(pyArr, i, j), stringLength);

        ExcelArrayBuilder builder(dims[0], dims[1], stringLength);
        for (auto i = 0; i < dims[0]; ++i)
          for (auto j = 0; j < dims[1]; ++j)
            TImpl::builderEmplace(builder, i, j, PyArray_GETPTR2(pyArr, i, j));

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