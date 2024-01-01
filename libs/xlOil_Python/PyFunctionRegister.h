#pragma once
#include "PyCore.h"
#include "TypeConversion/PyDictType.h"
#include <xlOil/Register.h>
#include <xlOil/Throw.h>
#include <xlOil/Interface.h>
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
    class IPyToExcel;
    class PyAddin;

    namespace FunctionRegistry
    {
      /// <summary>
      /// Adds the specified module to the specified context if the module
      /// has not already been read. If the module already exists, just 
      /// returns a reference to it.
      /// </summary>
      std::shared_ptr<RegisteredModule>
        addModule(
          const std::weak_ptr<PyAddin>& context,
          const std::wstring& modulePath,
          const wchar_t* workbookName);
    };

    struct PyFuncArg
    {
      PyFuncArg(
        std::wstring&& name, 
        std::wstring&& help, 
        const std::shared_ptr<IPyFromExcel>& converter,
        std::string&& flags);

      std::shared_ptr<IPyFromExcel> converter;
      std::wstring name;
      std::wstring help;
      pybind11::object default;
      std::string flags;

      bool isKeywords() const {
        return flags.find("keywords") != std::string::npos;
      }

      bool isVargs() const {
        return flags.find("vargs") != std::string::npos;
      }

      bool isArray() const {
        return flags.find("array") != std::string::npos;
      }
      std::string str() const;
    };

    class PyFuncInfo
    {
    public:
      PyFuncInfo(
        const pybind11::function& func,
        const std::vector<PyFuncArg>& args,
        const std::wstring& name,
        const std::string& features,
        const std::wstring& help,
        const std::wstring& category,
        bool isLocal,
        bool isVolatile,
        unsigned errorPropagation);

      ~PyFuncInfo();

      const auto& name() const { return _info->name; }

      auto& args() { return _args; }
      const auto& args() const { return _args; }

      auto getReturnConverter() const { return returnConverter; }
      void setReturnConverter(const std::shared_ptr<const IPyToExcel>& conv);
      auto getErrorPropagation() const { return (_propagateErrors & ADDIN) > 0; }
      void setErrorPropagation(bool x) { _propagateErrors |= x ? ADDIN : ~ADDIN; }

      bool isLocalFunc;
      bool isAsync;
      bool isRtdAsync;
      bool isThreadSafe() const { return (_info->options & FuncInfo::THREAD_SAFE) != 0; }
      bool isCommand()    const { return (_info->options & FuncInfo::COMMAND) != 0; }
      bool isFPArray()    const { return (_info->options & FuncInfo::ARRAY) != 0; }
      bool propagateErrors() const 
      { 
        auto x = (_propagateErrors & (ALWAYS | NEVER | ADDIN));
        return x > 0 && x < NEVER;
      }

      const std::shared_ptr<FuncInfo>& info() const { return _info; }

      const pybind11::function& func() const { return _func; }
      void setFunc(const pybind11::function& f) { _func = f; }

      static std::shared_ptr<const DynamicSpec> 
        createSpec(
          const std::shared_ptr<PyFuncInfo>& funcInfo,
          const PyAddin& addin);

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
        assert(pyArgs.capacity() >= _numPositionalArgs + (isRtdAsync || isAsync ? 1u : 0u));

        size_t i = 0;
        try
        {
          for (; i < _numPositionalArgs; ++i)
          {
            auto* defaultValue = _args[i].default.ptr();
            pyArgs.push_back((*_args[i].converter)(xlArgs(i), defaultValue));
          }

          if (_hasKeywordArgs)
          {
            if (!xlArgs(i).isMissing())
              kwargs = PySteal<>(readKeywordArgs(xlArgs(i)));
            ++i;
          }

          if (_hasVariableArgs)
          {
            auto& converter = *_args[i].converter;
            const auto* defaultValue = _args[i].default.ptr();
#ifdef _WIN64
            const auto maxArgs = XL_MAX_VBA_FUNCTION_ARGS - _args.size();
#else
            const auto maxArgs = 16 - _args.size();
#endif
            for (auto iVarg = i; iVarg < maxArgs && !xlArgs(iVarg).isMissing(); ++iVarg)
              pyArgs.push_back(converter(xlArgs(iVarg), defaultValue));
          }
        }
        catch (const std::exception& e)
        {
          // We give the arg number 1-based as it's more natural
          XLO_THROW(L"Error in arg {1} '{0}': {2}",
            _args[i].name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
        }
      }

      template<class TXlArgs>
      auto invoke(TXlArgs&& xlArgs) const
      {
        PyCallArgs<> pyArgs;
        py::object kwargs;

        convertArgs(
          std::forward<TXlArgs>(xlArgs),
          pyArgs,
          kwargs);

        return pyArgs.call(_func, kwargs);
      }

    private:
      std::shared_ptr<const IPyToExcel> returnConverter;
      std::vector<PyFuncArg> _args;
      std::shared_ptr<FuncInfo> _info;
      pybind11::function _func;
      bool _hasKeywordArgs;
      bool _hasVariableArgs;
      uint16_t _numPositionalArgs;
      unsigned _propagateErrors;

      void writeExcelArgumentDescription();

      enum Propagation
      {
        ALWAYS = 1 << 2,
        NEVER  = 1 << 3,
        ADDIN  = 1  // Set if error propagation is enabled by the addin
      };
    };

    class RegisteredModule : public LinkedSource
    {
    public:
      /// <summary>
      /// If provided, a linked workbook can be used for local functions
      /// </summary>
      /// <param name="modulePath"></param>
      /// <param name="workbookName"></param>
      RegisteredModule(
        const std::wstring& modulePath,
        const std::weak_ptr<PyAddin>& addin,
        const wchar_t* workbookName);

      ~RegisteredModule();

      void registerPyFuncs(
        const pybind11::handle& pyModule,
        const std::vector<std::shared_ptr<PyFuncInfo>>& functions,
        const bool append);

      void reload() override;

      void renameWorkbook(const wchar_t* newPathName) override;

    private:
      bool _linkedWorkbook;
      std::weak_ptr<PyAddin> _addin;
      pybind11::object _module;
    };
  }
}