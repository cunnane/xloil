#pragma once
#include "PyCore.h"
#include <xlOil/Register.h>
#include <map>
#include <string>
#include <pybind11/pybind11.h>

namespace xloil {
  class AddinContext; 
  struct FuncInfo; 
  class ExcelObj; 
  class DynamicSpec;
  template <class T> class IConvertToExcel;
}
namespace xloil 
{
  namespace Python
  {
    class RegisteredModule; 
    class IPyFromExcel; 
    using IPyToExcel = IConvertToExcel<PyObject>;

    namespace FunctionRegistry
    {
      /// <summary>
      /// Adds the specified module to the specified context if the module
      /// has not already been read. If the module already exists, just 
      /// returns a reference to it.
      /// </summary>
      std::shared_ptr<RegisteredModule>
        addModule(
          AddinContext* context, 
          const std::wstring& modulePath,
          const wchar_t* workbookName);
    };

    class PyFuncArg
    {
    private:
      std::shared_ptr<FuncInfo> _info;
      pybind11::object _default;

    public:
      PyFuncArg(std::shared_ptr<FuncInfo> info, unsigned i)
        : _info(info)
        , arg(_info->args[i])
      {}

      FuncArg& arg;

      std::shared_ptr<IPyFromExcel> converter;
      
      void setName(const std::wstring& value) { arg.name = value; }
      const auto& getName() const { return arg.name; }

      void setHelp(const std::wstring& value) { arg.help = value; }
      const auto& getHelp() const { return arg.help; }

      void setDefault(const pybind11::object& value) 
      {
        arg.type |= FuncArg::Optional;
        _default = value; 
      }
      const auto& getDefault() const 
      { 
        // what to return if this is null???
        return _default; 
      }
    };

    class PyFuncInfo
    {
    public:
      PyFuncInfo(
        const pybind11::function& func,
        const std::wstring& name,
        const unsigned numArgs,
        const std::string& features,
        const std::wstring& help,
        const std::wstring& category,
        bool isLocal,
        bool isVolatile,
        bool hasKeywordArgs);
      
      ~PyFuncInfo();

      auto& args() { return _args; }
      const auto& constArgs() const { return _args; }

      void setFuncOptions(unsigned val);

      auto getReturnConverter() const { return returnConverter; }
      void setReturnConverter(
        const std::shared_ptr<const IPyToExcel>& conv);

      void convertArgs(
        const ExcelObj** xlArgs, pybind11::object& args, pybind11::object& kwargs) const;

      void invoke(
        ExcelObj& result, PyObject* args, PyObject* kwargs) const noexcept;

      void invoke(
        PyObject* args, PyObject* kwargs) const;

      bool isLocalFunc;
      bool isAsync;
      bool isRtdAsync;
      bool isThreadSafe() const { return (_info->options & FuncInfo::THREAD_SAFE) != 0; }
      const std::shared_ptr<FuncInfo>& info() const { return _info; }

      static std::shared_ptr<const DynamicSpec> createSpec(const std::shared_ptr<const PyFuncInfo>& funcInfo);

    private:
      std::shared_ptr<const IPyToExcel> returnConverter;
      std::vector<PyFuncArg> _args;
      std::shared_ptr<FuncInfo> _info;
      pybind11::function _func;
      bool _hasKeywordArgs;
      size_t _numPositionalArgs;

      void checkArgConverters() const;
      pybind11::tuple convertArgs(const ExcelObj** xlArgs) const;
    };
  }
}