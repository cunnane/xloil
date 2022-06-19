#include "PyHelpers.h"
#include <xlOil/ExcelArray.h>
#include "BasicTypes.h"
#include "PyCore.h"
#include "PyExcelArrayType.h"
#include "PyEvents.h"
#include <pybind11/pybind11.h>

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    class UserImpl : public detail::PyFromAny
    {
      PyExcelArray* _ArrayWrapper = nullptr;

    public:
      using detail::PyFromAny::operator();

      PyObject* operator()(const ArrayVal& arr)
      {
        _ArrayWrapper = new PyExcelArray(arr);
        return py::cast(
          _ArrayWrapper,
          py::return_value_policy::take_ownership).release().ptr();
      }

      PyObject* operator()(const PStringRef& pstr)
      {
        return detail::PyFromString()(pstr);
      }

      void checkArrayWrapperDisposed()
      {
        if (_ArrayWrapper && _ArrayWrapper->refCount() != 1)
          XLO_THROW("Held reference to ExcelArray detected. Accessing this object in python may crash Excel");
      }

      constexpr wchar_t* failMessage() const { return L"Custom converter failed"; }
    };

    class CustomConverter : public IPyFromExcel
    {
    private:
      py::object _callable;
      bool _checkCache;

    public:
      CustomConverter(py::object&& callable, bool checkCache)
        : _callable(callable)
        , _checkCache(checkCache)
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
        PyFromExcel<UserImpl, TUseCache> typeConverter;
        auto arg = PySteal(typeConverter(xl, defaultVal));
        auto retVal = _callable(arg);
        typeConverter._impl.checkArrayWrapperDisposed();
        return retVal.release().ptr();
      }

      const char* name() const override
      {
        return _callable.ptr()->ob_type->tp_name;
      }
    };
 
    class CustomReturn : public IPyToExcel
    {
    private:
      py::object _callable;
    public:
      CustomReturn(py::object&& callable)
        : _callable(callable)
      {}
      virtual ~CustomReturn()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }
      virtual ExcelObj operator()(const PyObject& pyObj) const override
      {
        // Use raw C API for extra speed as this code is on a critical path
#if PY_VERSION_HEX < 0x03080000
        auto result = PyObject_CallFunctionObjArgs(_callable.ptr(), const_cast<PyObject*>(&pyObj), nullptr);
#elif PY_VERSION_HEX < 0x03090000
        PyObject* args[] = { nullptr, const_cast<PyObject*>(&pyObj) };
        auto result = _PyObject_Vectorcall(_callable.ptr(), args + 1, 1 | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#else
        auto result = PyObject_CallOneArg(_callable.ptr(), const_cast<PyObject*>(&pyObj));
#endif
        if (!result)
        {
          auto error = PyErr_Occurred();
          if (!PyErr_GivenExceptionMatches(error, cannotConvertException))
            throw py::error_already_set();
          PyErr_Clear();
          return ExcelObj();
        }
        auto converted = PySteal<>(result);
        // TODO: the custom converter should be able to specify the return type rather than generic FromPyObj
        // TODO: the user could create an infinite loop which cycles between two type converters - best way to avoid?
        return FromPyObj()(converted.ptr());
      }

      const py::object& handler() const { return _callable; }
    };

    static int theBinder = addBinder([](py::module& mod)
    {
      py::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>(mod, 
        "CustomConverter", R"(
          This is the interface class for custom type converters to allow them
          to be called from the Core.
        )")
        .def(py::init<py::object, bool>(), py::arg("callable"), py::arg("check_cache")=true);

      py::class_<CustomReturn, IPyToExcel, std::shared_ptr<CustomReturn>>(mod, 
        "CustomReturn")
        .def(py::init<py::object>(), py::arg("callable"))
        .def("get_handler", &CustomReturn::handler);
    });
  }
}