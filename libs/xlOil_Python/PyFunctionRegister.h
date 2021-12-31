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
          AddinContext& context, 
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
      /// <param name="args">Size must equal `argArraySize()`</param>
      /// <param name="kwargs"></param>
      void convertArgs(
        const ExcelObj** xlArgs,
        PyObject** args,
        pybind11::object& kwargs) const;

      template <class TArray>
      void convertArgs(
        const ExcelObj** xlArgs,
        TArray& args,
        pybind11::object& kwargsDict) const
      {
        assert(args.size() >= argArraySize());
        convertArgs(xlArgs, (PyObject**)args.data(), kwargsDict);
      }

      void invoke(
        ExcelObj& result,
        PyObject* const* args,
        PyObject* kwargsDict) const;

      template <class TArray>
      void invoke(
        ExcelObj& result,
        const TArray& args,
        PyObject* kwargsDict) const
      {
        assert(args.size() >= argArraySize());
        invoke(result, (PyObject* const*)args.data(), kwargsDict);
      }

      pybind11::object invoke(
        PyObject* const* args, const size_t nArgs, PyObject* kwargsDict) const;

      pybind11::object invoke(
        const std::vector<pybind11::object>& args, PyObject* kwargsDict) const
      {
        return invoke((PyObject* const*)args.data(), args.size() - theVectorCallOffset, kwargsDict);
      }

      /// <summary>
      /// Python can optimise onward calls to PyObject_Vectorcall if we
      /// leave a free entry at the start of the args array passed 
      /// through the API. For Py 3.7 and earlier vector call is not
      /// available so this is not required.  See <see cref="theVectorCallOffset"/>
      /// </summary>
      size_t argArraySize() const noexcept { return _numPositionalArgs + theVectorCallOffset; }

#if PY_VERSION_HEX < 0x03080000
      static constexpr auto theVectorCallOffset = 0u;
#else

      static constexpr auto theVectorCallOffset = 1u;
#endif

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