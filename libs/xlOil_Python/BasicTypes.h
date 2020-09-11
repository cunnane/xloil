#pragma once

#include "Numpy.h"
#include "Cache.h"
#include "Date.h"
#include "Main.h"
#include "Tuple.h"
#include "ExcelErrors.h"
#include "PyHelpers.h"
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/TypeConverters.h>
#include <xlOil/Log.h>
#include <xloil/StringUtils.h>
#include <xlOil/ExcelRange.h>
#include <string>

using namespace std::literals::string_literals;

namespace xloil 
{
  namespace Python
  {
    class IPyFromExcel : public IConvertFromExcel<PyObject*>
    {
    public:
      virtual PyObject* fromArray(const ExcelArray& arr) const = 0;
    };
    using IPyToExcel = IConvertToExcel<PyObject>;

    template<class TSuper=nullptr_t>
    class PyFromCache : public CacheConverter<PyObject*, NullCoerce<TSuper, PyFromCache<>>>
    {
    public:
      using base_type = CacheConverter;
      PyObject* fromString(const PStringView<>& pstr) const
      {
        pybind11::object cached;
        if (pyCacheGet(pstr.view(), cached))
          return cached.release().ptr();
        return base_type::fromString(pstr);
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
      PyObject * fromString(const PStringView<>& pstr) const
      {
        return PyUnicode_FromWideChar(const_cast<wchar_t*>(pstr.pstr()), pstr.length());
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
    class PyFromAny : public PyFromCache<NullCoerce<TSuper, PyFromAny<>>>
    {
    public:
      PyObject* fromInt(int x) const { return PyFromInt().fromInt(x); }
      PyObject* fromBool(bool x) const { return PyFromBool().fromBool(x); }
      PyObject* fromDouble(double x) const { return PyFromDouble().fromDouble(x); }
      PyObject* fromArray(const ExcelObj& obj) const { return excelArrayToNumpyArray(ExcelArray(obj)); }
      
      PyObject* fromEmpty(const PyObject*) const { Py_RETURN_NONE; }

      PyObject* fromString(const PStringView<>& pstr) const 
      { 
        auto result = PyFromCache<PyFromAny>::fromString(pstr);
        if (result)
          return result;
        return PyFromString().fromString(pstr);
      }

      PyObject * fromError(CellError err) const
      {
        auto pyObj = pybind11::cast(err);
        return pyObj.release().ptr();
      }
      PyObject * fromRef(const ExcelObj& obj) const
      {
        return pybind11::cast(newXllRange(obj)).release().ptr();
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

      PyObject* fromString(const PStringView<>& pstr) const
      {
        pybind11::object cached;
        if (pyCacheGet(pstr.view(), cached))
        {
          // Type checking seems nice, but it's unpythonic to raise an error here
          if (_typeCheck && PyObject_IsInstance(cached.ptr(), _typeCheck) == 0)
            XLO_WARN(L"Found `{0}` in cache but type was expected", pstr.string());
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

    struct FromPyLong
    {
      auto operator()(const PyObject* obj) const
      {
        if (!PyLong_Check(obj))
          XLO_THROW("Expected python int, got '{0}'", pyToStr(obj));
        return ExcelObj(PyLong_AsLong((PyObject*)obj));
      }
    };
    struct FromPyFloat
    {
      auto operator()(const PyObject* obj) const
      {
        if (!PyFloat_Check(obj))
          XLO_THROW("Expected python float, got '{0}'", pyToStr(obj));
        return ExcelObj(PyFloat_AS_DOUBLE(obj));
      }
    };
    struct FromPyBool
    {
      auto operator()(const PyObject* obj) const
      {
        if (!PyBool_Check(obj))
          XLO_THROW("Expected python bool, got '{0}'", pyToStr(obj));
        return ExcelObj(PyObject_IsTrue((PyObject*)obj) > 0);
      }
    };

    struct FromPyString
    {
      template <class TAlloc = PStringAllocator<wchar_t>>
      auto operator()(
        const PyObject* obj, 
        const TAlloc& allocator = PStringAllocator<wchar_t>()) const
      {
        if (!PyUnicode_Check(obj))
          XLO_THROW("Expected python str, got '{0}'", pyToStr(obj));

        const auto len = (char16_t)std::min<size_t>(
          USHRT_MAX, PyUnicode_GetLength((PyObject*)obj));
        PString<wchar_t, TAlloc> pstr(len, allocator);
        PyUnicode_AsWideChar((PyObject*)obj, pstr.pstr(), pstr.length());
        return ExcelObj(std::move(pstr));
      }
    };

    extern std::shared_ptr<const IPyToExcel> theCustomReturnConverter;

    struct FromPyObj
    {
      template <class TAlloc = PStringAllocator<wchar_t>>
      auto operator()(
        const PyObject* obj, 
        bool useCache = true, 
        const TAlloc& stringAllocator = PStringAllocator<wchar_t>()) const
      {
        auto p = (PyObject*)obj; // Python API isn't const-aware
        if (p == Py_None)
        {
          // Return #N/A here as xltypeNil is turned to zero by Excel
          return ExcelObj(CellError::NA);
        }
        else if (PyLong_Check(p))
        {
          return ExcelObj(PyLong_AsLong(p));
        }
        else if (PyFloat_Check(p))
        {
          return ExcelObj(PyFloat_AS_DOUBLE(p));
        }
        else if (PyBool_Check(p))
        {
          return ExcelObj(PyObject_IsTrue(p) > 0);
        }
        else if (isNumpyArray(p))
        {
          return ExcelObj(numpyArrayToExcel(p));
        }
        else if (isPyDate(p))
        {
          return ExcelObj(pyDateToExcel(p));
        }
        else if (Py_TYPE(p) == pyExcelErrorType)
        {
          auto err = pybind11::reinterpret_borrow<pybind11::object>(p).cast<CellError>();
          return ExcelObj(err);
        }
        else if (PyUnicode_Check(p))
        {
          return FromPyString()(p, stringAllocator);
        }
        else if (theCustomReturnConverter)
        {
          auto val = (*theCustomReturnConverter)(*p);
          if (!val.isType(ExcelType::Nil))
            return ExcelObj(std::move(val));
        }
        
        if (PyIterable_Check(p))
        {
          return nestedIterableToExcel(p);
        }
        else if (useCache)
        {
          return pyCacheAdd(PyBorrow<>(p));
        }
        else
          return ExcelObj(CellError::Value);
      }
    };
  }
}