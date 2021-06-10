#pragma once

#include "BasicTypes.h"
#include <pybind11/pybind11.h>
#include <functional>

namespace xloil
{
  namespace Python
  {
    constexpr char* const theInjectedModuleName = "xloil_core";
    constexpr char* const theReadConverterPrefix = "Read_";
    constexpr char* const theReturnConverterPrefix = "Return_";

    PyObject* buildInjectedModule();

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
    auto bindPyConverter(pybind11::module& mod, const char* type)
    {
      // TODO: static string concat?
      return pybind11::class_<T, IPyFromExcel, std::shared_ptr<T>>
        (mod, (theReadConverterPrefix + std::string(type)).c_str());
    }

    /// <summary>
    /// Declare a class of type IPyToExcel which handles the 
    /// specified type. Returns a reference to the bound class.
    /// </summary>
    template <class T>
    auto bindXlConverter(pybind11::module& mod, const char* type)
    {
      return pybind11::class_<T, IPyToExcel, std::shared_ptr<T>>(mod, 
        (theReturnConverterPrefix + std::string(type)).c_str());
    }
  }
}