#pragma once
#include <xlOil/TypeConverters.h>
#include <pybind11/pybind11.h>
#include <functional>

namespace xloil
{
  namespace Python
  {
    constexpr char* const theInjectedModuleName = "xloil_core";
    PyObject* buildInjectedModule();

    class IPyFromExcel : public IConvertFromExcel<PyObject*>
    {
    public:
      virtual PyObject* fromArray(const ExcelArray& arr) const = 0;
    };
    using IPyToExcel = IConvertToExcel<PyObject>;

    /// <summary>
    /// Registers a binder, that is, a function which binds types in the
    /// xlOil core moddule. This should be called from a static initialiser.
    /// Higher priority items are bound first, this allows coarse control
    /// over dependencies.
    /// </summary>
    int addBinder(
      std::function<void(pybind11::module&)> binder, 
      size_t priority = 0);

    /// <summary>
    /// Declare a class of type IPyFromExcel which handles the 
    /// specified type. Returns a reference to the bound class.
    /// </summary>
    template <class T>
    auto bindFrom(pybind11::module& mod, const char* type)
    {
      // TODO: static string concat?
      return pybind11::class_<T, IPyFromExcel, std::shared_ptr<T>>
        (mod, ("To_" + std::string(type)).c_str());
    }

    /// <summary>
    /// Declare a class of type IPyToExcel which handles the 
    /// specified type. Returns a reference to the bound class.
    /// </summary>
    template <class T>
    auto bindTo(pybind11::module& mod, const char* type)
    {
      return pybind11::class_<T, IPyToExcel, std::shared_ptr<T>>(mod, 
        ("From_" + std::string(type)).c_str());
    }

    /// <summary>
    /// Type object correponding to the bound xloil::CellError
    /// </summary>
    extern PyTypeObject* pyExcelErrorType;
  }
}