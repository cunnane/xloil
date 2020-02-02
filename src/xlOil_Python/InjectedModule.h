#pragma once
#include <pybind11/pybind11.h>
#include <functional>

// couldn't we use a global variable for this?

#define XLO_PY_MOD xloil_core
#define XLO_PY_MOD_STR "xloil_core"

namespace xloil
{
  namespace Python
  {
    PyObject* buildInjectedModule();
    int addBinder(std::function<void(pybind11::module&)> binder);

    template <class T>
    auto bindFrom(pybind11::module& mod, const char* type)
    {
      // TODO: static string concat?
      return pybind11::class_<T, IPyFromExcel, std::shared_ptr<T>>(mod, 
        (std::string(type) + "_from_Excel").c_str());
    }
    template <class T>
    auto bindTo(pybind11::module& mod, const char* type)
    {
      return pybind11::class_<T, IPyToExcel, std::shared_ptr<T>>(mod, 
        (std::string(type) + "_to_Excel").c_str());
    }

    void scanModule(pybind11::object& mod);
  }
}