#pragma once
#include <pybind11/pybind11.h>
#include <map>
#include <string>

namespace xloil { class AddinContext; }
namespace xloil 
{
  namespace Python
  {
    class RegisteredModule;

    namespace FunctionRegistry
    {
      std::shared_ptr<RegisteredModule>
        addModule(
          AddinContext* context, 
          const pybind11::module& moduleHandle, 
          const wchar_t* workbookName = nullptr);

      std::shared_ptr<RegisteredModule>
        addModule(
          AddinContext* context, 
          const std::wstring& modulePath, 
          const wchar_t* workbookName = nullptr);
    };
  }
}