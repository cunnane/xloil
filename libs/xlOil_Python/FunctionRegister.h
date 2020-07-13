#pragma once
#include "InjectedModule.h"
#include <map>
#include <string>
#include <pybind11/pybind11.h>

namespace xloil {
  class AddinContext; struct FuncInfo;  class ExcelObj; 
}
namespace xloil 
{
  namespace Python
  {
    class RegisteredModule;
    class IPyFromExcel;

    namespace FunctionRegistry
    {
      std::shared_ptr<RegisteredModule>
        addModule(
          AddinContext* context, 
          const std::wstring& modulePath, 
          const bool watchSource,
          const wchar_t* workbookName = nullptr);
    };

    class PyFuncInfo
    {
    public:
      PyFuncInfo(
        const std::shared_ptr<FuncInfo>& info,
        const pybind11::function& func,
        bool keywordArgs);

      void setArgTypeDefault(
        size_t i, 
        std::shared_ptr<IPyFromExcel> converter, 
        pybind11::object defaultVal);

      void setArgType(
        size_t i, 
        std::shared_ptr<IPyFromExcel> converter);

      void setFuncOptions(
        int val);

      std::pair<pybind11::tuple, pybind11::object> convertArgs(
        const ExcelObj** xlArgs);

      void invoke(
        ExcelObj& result, PyObject* args, PyObject* kwargs) noexcept;

      void invoke(
        PyObject* args, PyObject* kwargs);

      std::shared_ptr<FuncInfo> info;
      std::shared_ptr<IPyToExcel> returnConverter;
      bool isLocalFunc;
      bool isRtdAsync;

    private:
      pybind11::function func;
      std::vector<std::pair<std::shared_ptr<IPyFromExcel>, pybind11::object>> 
        argConverters;
      bool hasKeywordArgs;
    };
  }
}