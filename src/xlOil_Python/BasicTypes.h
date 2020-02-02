#pragma once

#include "ExcelObj.h"
#include "Numpy.h"
#include "xloil/Log.h"
#include "xloil/Utils.h"
#include "Cache.h"
#include "Date.h"
#include "Main.h"
#include "Tuple.h"
#include "InjectedModule.h"
#include "PyHelpers.h"
#include "TypeConverters.h"
#include <string>


using namespace std::literals::string_literals;

namespace xloil 
{
  namespace Python
  {
    using IPyFromExcel = IConvertFromExcel<PyObject*> ;
    using IPyToExcel = IConvertToExcel<PyObject> ;

    class PyFromDouble : public ConverterImpl<PyObject*>
    {
    public:
      PyObject * fromDouble(double x) const { return PyFloat_FromDouble(x); }
    };

    class PyFromBool : public ConverterImpl<PyObject*>
    {
    public:
      PyObject * fromBool(bool x) const { if (x) Py_RETURN_TRUE; else Py_RETURN_FALSE; }
    };

    class PyFromString : public ConverterImpl<PyObject*>
    {
    public:
      PyObject * fromString(const wchar_t* buf, size_t len) const
      {
        return PyUnicode_FromWideChar(const_cast<wchar_t*>(buf), len);
      }
      PyObject* fromEmpty(const PyObject*) const { return PyUnicode_New(0, 127); }
      PyObject* fromInt(int x) const { return PyUnicode_FromFormat("%i", x); }
      PyObject* fromBool(bool x) const { return PyUnicode_FromString(std::to_string(x).c_str()); }
      PyObject* fromDouble(double x) const { return PyUnicode_FromString(std::to_string(x).c_str()); }
    };

    class PyFromInt : public ConverterImpl<PyObject*>
    {
    public:
      PyObject* fromInt(int x) const { return PyLong_FromLong(long(x)); }
      PyObject* fromDouble(double x) const
      {
        long i;
        if (floatingToInt(x, i))
          return PyLong_FromLong(i);
        return nullptr;
      }
    };

    class PyFromAny : public ConverterImpl<PyObject*>
    {
    public:
      PyObject* fromInt(int x) const { return PyFromInt().fromInt(x); }
      PyObject* fromBool(bool x) const { return PyFromBool().fromBool(x); }
      PyObject* fromDouble(double x) const { return PyFromDouble().fromDouble(x); }
      PyObject* fromArray(const ExcelObj& obj) const { return excelArrayToNumpyArray2d(obj); }
      
      PyObject* fromError(CellError err) const;
      PyObject* fromEmpty(const PyObject*) const { Py_RETURN_NONE; }

      PyObject* fromString(const wchar_t* buf, size_t len) const 
      { 
        PyObject* cached = nullptr;
        if (theCore->maybeCacheReference(buf, len))
        {
          std::shared_ptr<const ExcelObj> obj;
          if (theCore->fetchCache(buf, len, obj))
            return FromExcel<PyFromAny>()(*obj);
        }
        else if (fetchCache(buf, len, cached))
          return cached;
        return PyFromString().fromString(buf, len); 
      }
    };
    
    class PyCacheObject : public ConverterImpl<PyObject*>
    {
      PyObject* _typeCheck = nullptr;
    public:
     // PyCacheObject(PyObject* typeCheck) : _typeCheck(typeCheck) {}

      PyObject* fromString(const wchar_t* buf, size_t len) const
      {
        PyObject* cached = nullptr;
        if (fetchCache(buf, len, cached))
        {
          // Type checking seems nice, but it's unpythonic to raise an error here
          if (_typeCheck && PyObject_IsInstance(cached, _typeCheck) == 0)
            XLO_WARN(L"Found `{0}` in cache but type was expected", std::wstring(buf, buf + len));
          return cached;
        }
        return nullptr;
      }
    };

    template <class TImpl>
    class CheckedFromExcel
    {
      FromExcel<TImpl> _impl;
    public:
      typedef PyObject* return_type;

      template <class...Args>
      CheckedFromExcel(Args&&...args) : _impl(std::forward<Args>(args)...) 
      {}
      return_type operator()(const ExcelObj& xl, const PyObject* defaultVal = nullptr) const
      {
        PyObject* ret = _impl(xl, defaultVal);
        if (!ret)
        {
          XLO_THROW(L"Failed converting "s + xl.toString() + L": "s
            + pyErrIfOccurred());
        }
        return ret;
      }
    };

    template <class TImpl>
    class PyFromExcel : public ConvertFromExcel<CheckedFromExcel<TImpl>>
    {
    public:
      template <class...Args>
      PyFromExcel(Args&&...args) 
        : ConvertFromExcel(std::forward<Args>(args)...) 
      {}
    };

    //template<class T>
    //class Defaulted : public T
    //{
    //  PyObject* _default;
    //public:
    //  template <class...Args>
    //  Defaulted(PyObject* val, Args&&...args) : T(args...), _default(val) {}
    //  PyObject* fromMissing(const PyObject*) const { Py_IncRef(_default); return _default; }
    //};

    inline ExcelObj fromPyLong(const PyObject* obj)
    {
      return ExcelObj(PyLong_AsLong((PyObject*)obj));
    }
    inline ExcelObj fromPyFloat(const PyObject* obj)
    {
      return ExcelObj(PyFloat_AS_DOUBLE(obj));
    }
    inline ExcelObj fromPyBool(const PyObject* obj)
    {
      return ExcelObj(PyObject_IsTrue((PyObject*)obj) > 0);
    }

    struct FromPyString
    {
      template <class TCtor>
      auto operator()(const PyObject* obj, TCtor ctor) const
      {
        size_t len = PyUnicode_GetLength((PyObject*)obj);
        wchar_t* buf;
        auto retVal = ctor(buf, len);
        PyUnicode_AsWideChar((PyObject*)obj, buf, len);
        return retVal;
      }
    };

    extern PyTypeObject* pyExcelErrorType;

    struct FromPyObj
    {
      template <class TCtor>
      auto operator()(const PyObject* obj, TCtor ctor) const
      {
        auto p = (PyObject*)obj; // Python API isn't const-aware
        if (p == Py_None)
        {
          // Return #N/A here as xltypeNil is turned to zero
          return ctor(CellError::NA);
        }
        else if (PyLong_Check(p))
        {
          return ctor(PyLong_AsLong(p));
        }
        else if (PyFloat_Check(p))
        {
          return ctor(PyFloat_AS_DOUBLE(p));
        }
        else if (PyBool_Check(p))
        {
          return ctor(PyObject_IsTrue(p) > 0);
        }
        else if (isNumpyArray(p))
        {
          return ctor(numpyArrayToExcel(p));
        }
        else if (isPyDate(p))
        {
          return ctor(pyDateToExcel(p));
        }
        else if (Py_TYPE(p) == pyExcelErrorType)
        {
          auto err = pybind11::reinterpret_borrow<pybind11::object>(p).cast<CellError>();
          return ctor(err);
        }
        else if (PyUnicode_Check(p))
        {
          return FromPyString()(p, ctor);
        }
        else if (PyIter_Check(p))
        {
          return ctor(nestedIterableToExcel(p));
        }
        else
        {
          return ctor(addCache(p));
        }

        //if (tmp = PySteal<decltype(tmp)>(PyObject_Str(p)))
      }
      auto operator()(const PyObject* obj) const
      {
        return operator()(obj, [](auto&&... args) { return ExcelObj(std::forward<decltype(args)>(args)...); });
      }
    };
  }
}