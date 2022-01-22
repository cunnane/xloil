#pragma once
#include "PyCore.h"
#include "TypeConversion/PyDictType.h"
#include <xlOil/Register.h>
#include <xlOil/Throw.h>
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
          AddinContext& context, 
          const std::wstring& modulePath,
          const wchar_t* workbookName);
    };

    class PyFuncArg
    {
    private:
      // We hold a ptr to the info as we take a FuncArg ref
      std::shared_ptr<FuncInfo> _info;
      pybind11::object _default;

    public:
      PyFuncArg(const std::shared_ptr<FuncInfo>& info, unsigned i)
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

      const auto& name() const { return _info->name; }

      auto& args() { return _args; }
      const auto& constArgs() const { return _args; }

      void setFuncOptions(unsigned val);

      auto getReturnConverter() const { return returnConverter; }
      void setReturnConverter(const std::shared_ptr<const IPyToExcel>& conv);

      bool isLocalFunc;
      bool isAsync;
      bool isRtdAsync;
      bool isThreadSafe() const { return (_info->options & FuncInfo::THREAD_SAFE) != 0; }
      bool isCommand()   const  { return (_info->options & FuncInfo::COMMAND) != 0; }

      const std::shared_ptr<FuncInfo>& info() const { return _info; }

      static std::shared_ptr<const DynamicSpec> createSpec(
        const std::shared_ptr<const PyFuncInfo>& funcInfo);
  
      /// <summary>
      /// Convert the array of ExcelObj arguments to PyObject values, with 
      /// option kwargs.
      /// </summary>
      /// <param name="xlArgs">Size must be equal to `args().size()`</param>
      /// <param name="pyArgs">Size must equal `argArraySize()`</param>
      /// <param name="kwargs"></param>
      template<class TXlArgs, class TPyArgs>
      void convertArgs(
        TXlArgs xlArgs,
        TPyArgs& pyArgs,
        pybind11::object& kwargs) const
      {
        assert(pyArgs.capacity() >= _numPositionalArgs + (isRtdAsync || isAsync ? 1 : 0));

        size_t i = 0;
        try
        {
          for (; i < _numPositionalArgs; ++i)
          {
            auto* defaultValue = _args[i].getDefault().ptr();
            pyArgs.push_back((*_args[i].converter)(xlArgs(i), defaultValue));
          }
          if (_hasKeywordArgs)
            kwargs = PySteal<>(readKeywordArgs(xlArgs(_numPositionalArgs)));
        }
        catch (const std::exception& e)
        {
          // We give the arg number 1-based as it's more natural
          XLO_THROW(L"Error in arg {1} '{0}': {2}",
            _args[i].arg.name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
        }
      }

      const pybind11::function& func() const { return _func; }

      ExcelObj convertReturn(PyObject* retVal) const;

      template<class TXlArgs> 
      ExcelObj invoke(TXlArgs&& xlArgs) const
      {
        PyCallArgs<> pyArgs;
        py::object kwargs;

        convertArgs(
          std::forward<TXlArgs>(xlArgs),
          pyArgs,
          kwargs);

        auto ret = PySteal<>(pyArgs.call(_func.ptr(), kwargs.ptr()));

        return std::move(convertReturn(ret.ptr()));
      }

    private:
      std::shared_ptr<const IPyToExcel> returnConverter;
      std::vector<PyFuncArg> _args;
      std::shared_ptr<FuncInfo> _info;
      pybind11::function _func;
      bool _hasKeywordArgs;
      uint16_t _numPositionalArgs;

      void checkArgConverters() const;
    };
  }
}