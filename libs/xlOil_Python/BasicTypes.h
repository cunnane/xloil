#pragma once

#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include "Numpy.h"
#include <xlOil/Log.h>
#include <xloilHelpers/StringUtils.h>
#include <xlOil/ExcelRange.h>
#include "Cache.h"
#include "Date.h"
#include "Main.h"
#include "Tuple.h"
#include "InjectedModule.h"
#include "PyHelpers.h"

#include <string>


using namespace std::literals::string_literals;

namespace xloil 
{
  namespace Python
  {
    template<class TSuper=nullptr_t>
    class PyFromCache : public CacheConverter<PyObject*, NotNull<TSuper, PyFromCache<>>>
    {
    public:
      using base_type = CacheConverter;
      PyObject* fromString(const wchar_t* buf, size_t len) const
      {
        pybind11::object cached;
        if (fetchCache(buf, len, cached))
          return cached.release().ptr();
        return base_type::fromString(buf, len);
      }
    };

    class PyFromDouble : public PyFromCache<PyFromDouble>
    {
    public:
      PyObject * fromDouble(double x) const { return PyFloat_FromDouble(x); }
    };

    class PyFromBool : public PyFromCache<PyFromBool>
    {
    public:
      PyObject * fromBool(bool x) const { if (x) Py_RETURN_TRUE; else Py_RETURN_FALSE; }
    };

    class PyFromString : public CacheConverter<PyObject*, PyFromString>
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

    class PyFromInt : public PyFromCache<PyFromInt>
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

    template<class TSuper = nullptr_t>
    class PyFromAny : public PyFromCache<NotNull<TSuper, PyFromAny<>>>
    {
    public:
      PyObject* fromInt(int x) const { return PyFromInt().fromInt(x); }
      PyObject* fromBool(bool x) const { return PyFromBool().fromBool(x); }
      PyObject* fromDouble(double x) const { return PyFromDouble().fromDouble(x); }
      PyObject* fromArray(const ExcelObj& obj) const { return excelArrayToNumpyArray(ExcelArray(obj)); }
      
      PyObject* fromEmpty(const PyObject*) const { Py_RETURN_NONE; }

      PyObject* fromString(const wchar_t* buf, size_t len) const 
      { 
        auto result = PyFromCache<PyFromAny>::fromString(buf, len);
        if (result)
          return result;
        return PyFromString().fromString(buf, len); 
      }

      PyObject * fromError(CellError err) const
      {
        auto pyObj = pybind11::cast(err);
        return pyObj.release().ptr();
      }
      PyObject * fromRef(const ExcelObj& obj) const
      {
        return pybind11::cast(ExcelRange(obj)).release().ptr();
      }
    };
    
    /// <summary>
    /// TODO: Not currently used but seems like a nice idea some time
    /// </summary>
    class PyCacheObject : public CacheConverter<PyObject*, PyCacheObject>
    {
      PyObject* _typeCheck = nullptr;
    public:
     // PyCacheObject(PyObject* typeCheck) : _typeCheck(typeCheck) {}

      PyObject* fromString(const wchar_t* buf, size_t len) const
      {
        pybind11::object cached;
        if (fetchCache(buf, len, cached))
        {
          // Type checking seems nice, but it's unpythonic to raise an error here
          if (_typeCheck && PyObject_IsInstance(cached.ptr(), _typeCheck) == 0)
            XLO_WARN(L"Found `{0}` in cache but type was expected", std::wstring(buf, buf + len));
          return cached.release().ptr();
        }
        return nullptr;
      }
    };

    template <class TImpl>
    class FromExcel
    {
      TImpl _impl;

    public:
      template <class...Args>
      FromExcel(Args&&...args) : _impl(std::forward<Args>(args)...) 
      {}

      PyObject* operator()(const ExcelObj& xl, const PyObject* defaultVal = nullptr) const
      {
        auto ret = _impl(xl, defaultVal);
        if (!ret)
          XLO_THROW(L"Failed converting {0}: {1}", xl.toString(), pyErrIfOccurred());
        
        return ret;
      }
      PyObject* fromArray(const ExcelArray& arr) const
      {
        auto ret = _impl.fromArrayObj(arr);
        if (!ret)
          XLO_THROW(L"Failed converting to array: {0}", pyErrIfOccurred());
        
        return ret;
      }
      TImpl& impl() const { return _impl._impl; }
    };

    template <class TImpl>
    class PyFromExcel : public IPyFromExcel
    {
      FromExcel<TImpl> _impl;
    public:
      template <class...Args>
      PyFromExcel(Args&&...args) 
        : _impl(std::forward<Args>(args)...)
      {}
      virtual PyObject* fromArray(const ExcelArray& arr) const override
      {
        return _impl.fromArray(arr);
      }
      virtual PyObject* operator()(const ExcelObj& xl, const PyObject* defaultVal = nullptr) const override
      {
        return _impl(xl, defaultVal);
      }
    };

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
        auto len = (char16_t)std::min<size_t>(USHRT_MAX, PyUnicode_GetLength((PyObject*)obj));
        PString<> pstr(len);
        PyUnicode_AsWideChar((PyObject*)obj, pstr.pstr(), pstr.length());
        return ctor(pstr);
      }
    };

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
        else if (PyIterable_Check(p))
        {
          return ctor(nestedIterableToExcel(p));
        }
        else
        {
          return ctor(addCache(PyBorrow<pybind11::object>(p)));
        }
      }
      auto operator()(const PyObject* obj) const
      {
        return operator()(obj, [](auto&&... args) { return ExcelObj(std::forward<decltype(args)>(args)...); });
      }
    };
  }
}