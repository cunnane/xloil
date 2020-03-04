#pragma once
#include <pybind11/pybind11.h>
#include <map>
#include <string>

namespace xloil 
{
  namespace Python
  {
    class RegisteredModule;
    struct PyFuncInfo;

    class FunctionRegistry
    {
    public:
      static FunctionRegistry& get();

      std::shared_ptr<RegisteredModule> addModule(const pybind11::module& moduleHandle);
      std::shared_ptr<RegisteredModule> addModule(const std::wstring& modulePath);

      auto & modules() { return _modules; }

    private:
      FunctionRegistry();
      std::map<std::wstring, std::shared_ptr<RegisteredModule>> _modules;
    };
  }
}