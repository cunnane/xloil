#include "PyHelpers.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "PyCore.h"
#include "PyExcelArrayType.h"
#include "PyEvents.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;
using std::string;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      /// <summary>
      /// Wraps the usual PyFromAny but intercepts the array handling to return 
      /// a PyExcelArray object, then checks that the object passed to python was
      /// properly disposed of and no reference to the temporary array object remains.
      /// </summary>
      class CustomConverterArrayHandler : 
        public detail::PyFromAny<CustomConverterArrayHandler>
      {
        py::object _ArrayWrapper;
        PyExcelArray* _excelArray;

      public:
        CustomConverterArrayHandler() : _excelArray(nullptr)
        {}

        using detail::PyFromAny<CustomConverterArrayHandler>::operator();

        PyObject* operator()(const ArrayVal& arr) const
        {
          const_cast<CustomConverterArrayHandler*>(this)->setWrapper(
            new PyExcelArray(arr));
          // Inc ref count as this function should return a stealable reference
          _ArrayWrapper.inc_ref();
          return _ArrayWrapper.ptr();
        }

        // Normally #N/A inputs are converted to None in PyFromAny. To give the custom converter
        // maximum flexibility we override that behaviour

        PyObject* operator()(CellError err) const
        {
          auto pyObj = pybind11::cast(err);
          return pyObj.release().ptr();
        }

        ~CustomConverterArrayHandler()
        {
          if (_excelArray && (_ArrayWrapper.ref_count() != 1 || _excelArray->refCount() != 1))
            XLO_ERROR("Held reference to ExcelArray detected. Accessing this object "
              "in python may crash Excel");
        }

        void setWrapper(PyExcelArray* ptr)
        {
          // This class should only be used for a single ExcelObj, then disposed of.
          assert(!_excelArray);
          _excelArray = ptr;
          _ArrayWrapper = py::cast(
            ptr,
            py::return_value_policy::take_ownership);
        }

        constexpr wchar_t* failMessage() const { return L"Custom converter failed"; }
      };
    }

    class CustomConverter : public IPyFromExcel
    {
    private:
      py::object _callable;
      bool _checkCache;
      string _name;

    public:
      CustomConverter(py::object&& callable, bool checkCache, const char* name)
        : _callable(callable)
        , _checkCache(checkCache)
        , _name(name)
      {}

      virtual ~CustomConverter()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }

      virtual result_type operator()(
        const ExcelObj& xl, 
        const_result_ptr defaultVal) override
      {
        // Called from type conversion code where GIL has already been acquired
        return checkUserException([&]()
        {
          if (_checkCache)
            return callConverter<true>(xl, defaultVal);
          else
            return callConverter<false>(xl, defaultVal);
        });
      }

      template<bool TUseCache>
      auto callConverter(const ExcelObj& xl, const_result_ptr defaultVal)
      {
        PyFromExcel<CustomConverterArrayHandler, TUseCache> typeConverter;
        auto arg = PySteal(typeConverter(xl, defaultVal));
        auto retVal = _callable(arg);
        return retVal.release().ptr();
      }

      const char* name() const override
      {
        return _name.c_str();
      }
    };
 
    class CustomReturn : public IPyToExcel
    {
    private:
      py::object _callable;
      string _name;
    public:
      CustomReturn(py::object&& callable, const char* name)
        : _callable(callable)
        , _name(name)
      {
      }
      virtual ~CustomReturn()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }
      virtual ExcelObj operator()(const PyObject* target) const override
      {
        auto* result = invokeImpl(target);
        if (!result)
        {
          auto error = PyErr_Occurred();
          if (!PyErr_GivenExceptionMatches(error, cannotConvertException))
            throw py::error_already_set();
          PyErr_Clear();
          return ExcelObj();
        }
        auto converted = PySteal(result);
        // TODO: the custom converter should be able to specify the return type rather than generic FromPyObj
        // TODO: the user could create an infinite loop which cycles between two type converters - best way to avoid?
        return FromPyObj()(converted.ptr());
      }
      auto invoke(const py::object& target) const
      {
        return PySteal(invokeImpl(target.ptr()));
      }

      PyObject* invokeImpl(const PyObject* target) const
      {
      // Use raw C API for extra speed as this code is on a critical path
#if PY_VERSION_HEX < 0x03080000
        auto result = PyObject_CallFunctionObjArgs(_callable.ptr(), const_cast<PyObject*>(target), nullptr);
#elif PY_VERSION_HEX < 0x03090000
        PyObject* args[] = { nullptr, const_cast<PyObject*>(target) };
        auto result = _PyObject_Vectorcall(_callable.ptr(), args + 1, 1 | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#else
        auto result = PyObject_CallOneArg(_callable.ptr(), const_cast<PyObject*>(target));
#endif
        return result;
      }

      const char* name() const override
      {
        return _name.c_str();
      }
    };

    static int theBinder = addBinder([](py::module& mod)
    {
      py::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>(mod, 
        "_CustomConverter", R"(
          This is the interface class for custom type converters to allow them
          to be called from the Core.
        )")
        .def(py::init<py::object, bool, const char*>(),
          py::arg("callable"), 
          py::arg("check_cache")=true,
          py::arg("name")="custom");

      py::class_<CustomReturn, IPyToExcel, std::shared_ptr<CustomReturn>>(mod, 
        "_CustomReturn")
        .def(py::init<py::object, const char*>(), py::arg("callable"), py::arg("name")="custom")
        .def("invoke", &CustomReturn::invoke);
    });
  }
}