#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include "Numpy.h"
#include "BasicTypes.h"
#include "TypeConverters.h"
#include "StandardConverters.h"
#include "ExcelArray.h"
#include "ArrayHelpers.h"
#include <numpy/arrayobject.h>
#include <numpy/arrayscalars.h>
#include <numpy/npy_math.h>
#include <pybind11/pybind11.h>

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

    template <class TConv, class TDataType, int TNpType>
    class PyFromArray1d : public PyFromCache<PyFromArray1d<TConv, TDataType, TNpType>>
    {
      bool _trim;
      TConv _conv;
    public:
      PyFromArray1d(bool trim) : _trim(trim) {}

      PyObject* fromArray(const ExcelObj& obj) const
      {
        ExcelArray arr(obj, _trim);
        return fromArray(arr);
      }
      PyObject* fromArray(const ExcelArray& arr) const
      {
        if (arr.size() == 0)
          Py_RETURN_NONE;

        if (arr.dims() != 1)
          XLO_THROW("Expecting a 1-dim array");

        Py_intptr_t dims[1];
        dims[0] = arr.size();
        const int nDims = 1;
        auto pyObj = PySteal<py::object>(PyArray_SimpleNew(nDims, dims, TNpType));
        auto pyArr = (PyArrayObject*)pyObj.ptr();
        for (auto i = 0; i < arr.size(); ++i)
        {
          TDataType val = _conv(arr(i));
          auto ptr = (TDataType*)PyArray_GETPTR1(pyArr, i);
          *ptr = val;
        }
        return pyObj.release().ptr();
      }
    };

    template <class TConv, class TDataType, int TNpType>
    class PyFromArray2d : public PyFromCache<PyFromArray2d<TConv, TDataType, TNpType>>
    {
      bool _trim;
      TConv _conv;
    public:
      PyFromArray2d(bool trim) : _trim(trim) {}

      PyObject* fromArray(const ExcelObj& obj) const
      {
        ExcelArray arr(obj, _trim);
        return fromArray(arr);
      }

      PyObject* fromArray(const ExcelArray& arr) const
      {
        if (arr.size() == 0)
          Py_RETURN_NONE;

        Py_intptr_t dims[2];
        const int nDims = 2;
        dims[0] = arr.nRows();
        dims[1] = arr.nCols();

        auto pyObj = PySteal<py::object>(PyArray_SimpleNew(nDims, dims, TNpType));
        auto pyArr = (PyArrayObject*)pyObj.ptr();
        for (auto i = 0; i < arr.nRows(); ++i)
          for (auto j = 0; j < arr.nCols(); ++j)
          {
            TDataType val = _conv(arr(i, j));
            auto ptr = (TDataType*)PyArray_GETPTR2(pyArr, i, j);
            *ptr = val;
            //PyArray_SETITEM(pyArr, (char*)PyArray_GETPTR2(pyArr, i, j), p);
          }
        return pyObj.release().ptr();
      }
    };

    template<class TConv>
    struct FromArrayPyObjectImpl
    {
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

    template <class TDataType, int TNpType>
    struct FromArrayDTypeImpl
    {
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

    template <class TImpl>
    class XlFromArray1d : public IConvertToExcel<PyObject>
    {
      bool _cache;

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

        ExcelArrayBuilder builder(1, dims[0]);
        for (auto j = 0; j < dims[0]; ++j)
          TImpl::builderEmplace(builder, 0, j, PyArray_GETPTR1(pyArr, j));
        
        return _cache
          ? theCore->insertCache(builder.toExcelObj())
          : builder.toExcelObj();
      }
    };

    template <class TImpl>
    class XlFromArray2d : public IConvertToExcel<PyObject>
    {
      bool _cache;
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

  
    PyObject* excelArrayToNumpyArray2d(const ExcelObj& obj)
    {
      ExcelArray arr(obj);
      return PyFromArray2d<CheckedFromExcel<PyFromAny>, PyObject*, NPY_OBJECT>(true).fromArray(arr);
    }

    template <template <class> class TThing>
    auto selectDataType(const PyArrayObject* pArr)
    {
      auto dType = PyArray_TYPE(pArr);
      auto& p = *(PyObject*)pArr;

      // TODO: do some non-ugly casting to implement commented dtypes
      switch (dType)
      {
      case NPY_BOOL: return TThing<FromArrayDTypeImpl<bool, NPY_BOOL>>()(p);
      //case NPY_BYTE: return TThing<FromArrayTypeImpl<char, NPY_BYTE>>()(p);
      //case NPY_UBYTE: return TThing<FromArrayTypeImpl<unsigned char, NPY_UBYTE>>()(p);
      case NPY_SHORT: return TThing<FromArrayDTypeImpl<short, NPY_SHORT>>()(p);
      case NPY_USHORT: return TThing<FromArrayDTypeImpl<unsigned short, NPY_USHORT>>()(p);
      case NPY_UINT:
      case NPY_INT: return TThing<FromArrayDTypeImpl<int, NPY_INT>>()(p);
      case NPY_LONG:
      case NPY_ULONG: return TThing<FromArrayDTypeImpl<long, NPY_ULONG>>()(p);
      //case NPY_LONGLONG: 
      //case NPY_ULONGLONG: return TThing<FromArrayTypeImpl<long long, NPY_ULONGLONG>>()(p);

      case NPY_FLOAT: return TThing<FromArrayDTypeImpl<float, NPY_FLOAT>>()(p);
      case NPY_DOUBLE: return TThing<FromArrayDTypeImpl<double, NPY_DOUBLE>>()(p);
      //case NPY_LONGDOUBLE: return TThing<FromArrayTypeImpl<long double, NPY_LONGDOUBLE>>()(p);

      // ????? case NPY_DATETIME:
      case NPY_OBJECT: return TThing<FromArrayPyObjectImpl<FromPyObj>>()(p);

      case NPY_STRING: 
      case NPY_UNICODE: return TThing<FromArrayPyObjectImpl<FromPyString>>()(p);
      default:
        XLO_THROW("Unsupported numpy date type");
      }
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
        return selectDataType<XlFromArray1d>(pArr);
      case 2:
        return selectDataType<XlFromArray2d>(pArr);
      default:
        XLO_THROW("Expected 1 or 2 dim array");
      }
    }

    namespace
    {
      template <class TConv, class TDataType, int TNpType>
      using Array1dFromXL = PyFromExcel<PyFromArray1d<TConv, TDataType, TNpType>>;

      template <class TConv, class TDataType, int TNpType>
      using Array2dFromXL = PyFromExcel<PyFromArray2d<TConv, TDataType, TNpType>>;

      template<class T>
      void declare(pybind11::module& mod, const char* name)
      {
        py::class_<T, IPyFromExcel, shared_ptr<T>>(mod, name)
          .def(py::init<bool>(), py::arg("trim")=true);
      }

      template<class T>
      void declare2(pybind11::module& mod, const char* name)
      {
        py::class_<T, IPyToExcel, shared_ptr<T>>(mod, name)
          .def(py::init<bool>(), py::arg("cache") = false);
      }
      static int theBinder = addBinder([](py::module& mod)
      {
        declare<Array1dFromXL<FromExcel<ToInt>, int, NPY_INT>                        >(mod, "Array_int_1d_from_Excel");
        declare<Array1dFromXL<FromExcel<ToDoubleNPYNan>, double, NPY_DOUBLE>         >(mod, "Array_float_1d_from_Excel");
        declare<Array1dFromXL<FromExcel<ToBool>, bool, NPY_BOOL>                     >(mod, "Array_bool_1d_from_Excel");
        declare<Array1dFromXL<CheckedFromExcel<PyFromString>, PyObject*, NPY_STRING> >(mod, "Array_str_1d_from_Excel");
        declare<Array1dFromXL<CheckedFromExcel<PyFromAny>, PyObject*, NPY_OBJECT>    >(mod, "Array_object_1d_from_Excel");

        declare<Array2dFromXL<FromExcel<ToInt>, int, NPY_INT>                        >(mod, "Array_int_2d_from_Excel");
        declare<Array2dFromXL<FromExcel<ToDoubleNPYNan>, double, NPY_DOUBLE>         >(mod, "Array_float_2d_from_Excel");
        declare<Array2dFromXL<FromExcel<ToBool>, bool, NPY_BOOL>                     >(mod, "Array_bool_2d_from_Excel");
        declare<Array2dFromXL<CheckedFromExcel<PyFromString>, PyObject*, NPY_STRING> >(mod, "Array_str_2d_from_Excel");
        declare<Array2dFromXL<CheckedFromExcel<PyFromAny>, PyObject*, NPY_OBJECT>    >(mod, "Array_object_2d_from_Excel");

        //declare2<XlFromArray1d<EmplacePyObj<FromPyObj>, PyObject*, NPY_OBJECT > > (mod, "Array_object_1d_to_Excel");
        //declare2<XlFromArray1d<EmplaceType<double>, double, NPY_DOUBLE > >(mod, "Array_float_1d_to_Excel");
      });
    }
  }
}