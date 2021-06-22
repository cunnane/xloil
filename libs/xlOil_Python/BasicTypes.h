#pragma once

#include "Numpy.h"
#include "Cache.h"
#include "Date.h"
#include "Main.h"
#include "Tuple.h"
#include "ExcelErrorType.h"
#include "PyHelpers.h"
#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/TypeConverters.h>
#include <xlOil/ExcelObjCache.h>
#include <xlOil/Log.h>
#include <xloil/StringUtils.h>
#include <xlOil/ExcelRange.h>
#include <string>

using namespace std::literals::string_literals;

namespace xloil 
{
  namespace Python
  {
    // TODO: rename this to IPyFromExcel, IPyToExcel
    using IPyFromExcel = IConvertFromExcel<PyObject*>;
    using IPyToExcel = IConvertToExcel<PyObject>;

    struct PyFromExcelImpl : public FromExcelBase<PyObject*>
    {
    };


    namespace detail
    {
      /// <summary>
      /// Wraps a type conversion functor, interpreting the string conversion to
      /// look for a python cache reference.  If found, returns the cache object,
      /// otherwise passes the string through.
      /// </summary>
      template<class TBase>
      struct PyFromCache : public TBase
      {
        template <class...Args>
        PyFromCache(Args&&...args) 
          : TBase(std::forward<Args>(args)...)
        {}

        using TBase::operator();
        PyObject* operator()(const PStringView<>& str) const
        {
          pybind11::object cached;
          if (pyCacheGet(str.view(), cached))
            return cached.release().ptr();
          return TBase::operator()(str);
        }
      };

      struct PyFromDouble : public PyFromExcelImpl
      {
        using PyFromExcelImpl::operator();
        PyObject* operator()(double x) const { return PyFloat_FromDouble(x); }
        PyObject* operator()(int x) const    { return operator()(double(x)); }
        PyObject* operator()(bool x) const   { return operator()(double(x)); }
        constexpr wchar_t* failMessage() const { return L"Expected float"; }
      };

      struct PyFromBool : public PyFromExcelImpl
      {
        using PyFromExcelImpl::operator();
        PyObject* operator()(bool x) const
        {
          if (x) Py_RETURN_TRUE; else Py_RETURN_FALSE;
        }
        PyObject* operator()(int x) const     { return operator()(bool(x)); }
        PyObject* operator()(double x) const  { return operator()(x != 0); }
        constexpr wchar_t* failMessage() const { return L"Expected bool"; }
      };

      struct PyFromString : public PyFromExcelImpl
      {
        using PyFromExcelImpl::operator();
        PyObject* operator()(const PStringView<>& pstr) const
        {
          return PyUnicode_FromWideChar(const_cast<wchar_t*>(pstr.pstr()), pstr.length());
        }
        // Return empty string for Excel Nil value
        PyObject* operator()(nullptr_t) const { return PyUnicode_New(0, 127); }
        PyObject* operator()(int x) const { return PyUnicode_FromFormat("%i", x); }
        PyObject* operator()(bool x) const { return PyUnicode_FromString(std::to_string(x).c_str()); }
        PyObject* operator()(double x) const { return PyUnicode_FromString(std::to_string(x).c_str()); }

        constexpr wchar_t* failMessage() const { return L"Expected string"; }
      };

      struct PyFromInt : public PyFromExcelImpl
      {
        using PyFromExcelImpl::operator();
        PyObject* operator()(int x) const  { return PyLong_FromLong(long(x)); }
        PyObject* operator()(bool x) const { return operator()(int(x)); }
        PyObject* operator()(double x) const
        {
          long i;
          if (floatingToInt(x, i))
            return PyLong_FromLong(i);
          return nullptr;
        }
        
        constexpr wchar_t* failMessage() const { return L"Expected int"; }
      };

      struct PyFromAny : public PyFromExcelImpl
      {
        using PyFromExcelImpl::operator();
        PyObject* operator()(int x) const { return PyFromInt()(x); }
        PyObject* operator()(bool x) const { return PyFromBool()(x); }
        PyObject* operator()(double x) const { return PyFromDouble()(x); }
        PyObject* operator()(ArrayVal arr) const
        {
          return excelArrayToNumpyArray(ExcelArray(arr));
        }

        // Return python None for Excel nil value
        PyObject* operator()(nullptr_t) const { Py_RETURN_NONE; }

        PyObject* operator()(const PStringView<>& pstr) const
        {
          return PyFromString()(pstr);
        }

        PyObject* operator()(CellError err) const
        {
          auto pyObj = pybind11::cast(err);
          return pyObj.release().ptr();
        }
        PyObject* operator()(RefVal ref) const
        {
          return pybind11::cast(newXllRange(ref)).release().ptr();
        }

        constexpr wchar_t* failMessage() const { return L"Unknown type"; }
      };

      /// <summary>
      /// Type converter which expects a cache reference string and rejects
      /// all other types.
      /// </summary>
      class PyCacheObject : public FromExcelBase<PyObject*>
      {
      public:
        using FromExcelBase::operator();

        PyObject* operator()(const PStringView<>& pstr) const
        {
          PyObject* _typeCheck = nullptr;

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

        constexpr wchar_t* failMessage() const { return L"Expected cache string"; }
      };
    }
    
    /// <summary>
    /// Wraps a type conversion implementation, similarly to <see cref="xloil::FromExcel"/>
    /// Checks all ExcelObj strings for both python and ExcelObj cache references.
    /// Throws an error if conversion fails. 
    /// </summary>
    template<class TImpl>
    struct PyFromExcel
    {
      detail::PyFromCache<CacheConverter<TImpl>> _impl;

      template <class...Args>
      PyFromExcel(Args&&...args)
        : _impl(std::forward<Args>(args)...)
      {}

      auto operator()(
        const ExcelObj& xl,
        const PyObject* defaultVal) const
      {
        return operator()(xl, const_cast<PyObject*>(defaultVal));
      }

      /// <summary>
      /// <returns>New/borrowed reference</returns>
      /// </summary>
      auto operator()(
        const ExcelObj& xl,
        PyObject* defaultVal = nullptr) const
      {
        if (xl.isMissing() && defaultVal)
        {
          // If we return the default value, we need to increment its refcount
          Py_INCREF(defaultVal);
          return defaultVal;
        }

        // Why return null and not throw here?
        auto* retVal = visitExcelObj(xl, _impl);

        if (!retVal)
        {
          XLO_THROW(L"Cannot convert {0}: {1}", xl.toString(),
            PyErr_Occurred() ? pyErrIfOccurred() : _impl.failMessage());
        }
        return retVal;
      }
    };

    using PyFromInt = PyFromExcel<detail::PyFromInt>;
    using PyFromBool = PyFromExcel<detail::PyFromBool>;
    using PyFromDouble = PyFromExcel<detail::PyFromDouble>;
    using PyFromString = PyFromExcel<detail::PyFromString>;
    using PyFromAny = PyFromExcel<detail::PyFromAny>;
    using PyCacheObject = PyFromExcel<detail::PyCacheObject>;

    namespace detail
    {
      /// <summary>
      /// Used by PyExcelConverter
      /// </summary>
      template <class T>
      struct MakePyFromExcel { using type = PyFromExcel<T>; };
      template <class T>
      struct MakePyFromExcel<PyFromExcel<T>> { using type = PyFromExcel<T>; };
    }

    /// <summary>
    /// Wraps a <see cref="PyFromExcel"/> to inherit from <see cref="IPyFromExcel"/>
    /// and create a type converter object with a virtual call.
    /// </summary>
    template <class TImpl>
    class PyExcelConverter : public IPyFromExcel
    {
      typename detail::MakePyFromExcel<TImpl>::type _impl;

    public:
      template <class...Args>
      PyExcelConverter(Args&&...args) 
        : _impl(std::forward<Args>(args)...)
      {}

      virtual PyObject* operator()(
        const ExcelObj& xl, const PyObject* defaultVal = nullptr) const override
      {
        // Because ref-counting there's no notion of a const PyObject*
        // for a default value
        return _impl(xl, const_cast<PyObject*>(defaultVal));
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
          USHRT_MAX, PyUnicode_GET_LENGTH((PyObject*)obj));
        PString<wchar_t, TAlloc> pstr(len, allocator);
        PyUnicode_AsWideChar((PyObject*)obj, pstr.pstr(), pstr.length());
        return ExcelObj(std::move(pstr));
      }
    };

    extern std::shared_ptr<const IPyToExcel> theCustomReturnConverter;

    template<bool TUseCache = true, CellError TFailure=CellError::Value>
    struct FromPyObj
    {
      template <class TAlloc = PStringAllocator<wchar_t>>
      auto operator()(
        const PyObject* obj, 
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
        else if (TUseCache)
        {
          return pyCacheAdd(PyBorrow<>(p));
        }
        else
          return ExcelObj(TFailure);
      }
    };
  }
}