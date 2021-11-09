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
    public:
      using detail::PyFromAny::operator();

      PyObject* operator()(ArrayVal arr) const
      {
        return py::cast(PyExcelArray(arr)).release().ptr();
      }
      // Override cache ref lookup?
      PyObject* operator()(const PStringView<>& pstr) const
      {
        return detail::PyFromString()(pstr);
      }
      constexpr wchar_t* failMessage() const { return L"Custom converter failed"; }
    };

    class CustomConverter : public IPyFromExcel
    {
    private:
      py::object _callable;
    public:
      CustomConverter(py::object&& callable)
        : _callable(callable)
      {}
      virtual ~CustomConverter()
      {
        py::gil_scoped_acquire getGil;
        _callable = py::object();
      }
      virtual result_type operator()(const ExcelObj& xl, const_result_ptr defaultVal) const override
      {
        // Called from type conversion code where GIL has already been acquired
        return checkUserException([&]() 
        {
          auto arg = PySteal(PyFromExcel<UserImpl>()(xl, defaultVal));
          return _callable(arg).release().ptr();
        });
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
        // This c
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
      py::class_<CustomConverter, IPyFromExcel, std::shared_ptr<CustomConverter>>
        (mod, "CustomConverter")
        .def(py::init<py::object>(), py::arg("callable"));

      py::class_<CustomReturn, IPyToExcel, std::shared_ptr<CustomReturn>>
        (mod, "CustomReturn")
        .def(py::init<py::object>(), py::arg("callable"))
        .def("get_handler", &CustomReturn::handler);

    });
  }
}