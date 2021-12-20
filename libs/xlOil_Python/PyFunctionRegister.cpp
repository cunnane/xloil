#include "PyFunctionRegister.h"
#include "PyCore.h"
#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "TypeConversion/PyDictType.h"
#include "PySource.h"
#include "AsyncFunctions.h"
#include "PyEvents.h"
#include "EventLoop.h"
#include <xloil/StaticRegister.h>
#include <xloil/DynamicRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Caller.h>
#include <xloil/RtdServer.h>
#include <xlOil/ExcelApp.h>
#include <xlOil/Interface.h>
#include <pybind11/stl.h>

#include <map>
#include <filesystem>

namespace fs = std::filesystem;
using std::shared_ptr;
using std::vector;
using std::pair;
using std::map;
using std::wstring;
using std::string;
using std::make_shared;
using std::make_pair;
using std::unique_ptr;
namespace py = pybind11;
using namespace pybind11::literals;

namespace xloil
{
  namespace Python
  {
    constexpr wchar_t* XLOPY_ANON_SOURCE = L"PythonFuncs";
    constexpr char* XLOPY_CLEANUP_FUNCTION = "_xloil_unload";

    void setFuncType(PyFuncInfo& info, const string& features, bool isVolatile)
    {
      unsigned base = isVolatile ? FuncInfo::VOLATILE : 0;
      if (features.empty())
        return;
      else if(features == "macro")
      {
        info.setFuncOptions(FuncInfo::MACRO_TYPE | base);
      }
      else if (features == "threaded")
      {
        info.setFuncOptions(FuncInfo::THREAD_SAFE | base);
        // turn off local?
      }
      else if (features == "rtd")
      {
        info.isRtdAsync = true;
      }
      else if (features == "async")
      {
        info.isAsync = true;
        // turn off local!
      }
      else
        throw py::value_error(formatStr("FuncSpec: Unknown function features '%s'", features.c_str()));
    }

    PyFuncInfo::PyFuncInfo(
      const pybind11::function& func,
      const std::wstring& name,
      const unsigned numArgs,
      const std::string& features,
      const std::wstring& help,
      const std::wstring& category,
      bool isLocal,
      bool isVolatile,
      bool hasKeywordArgs)
      : _info(new FuncInfo())
      , _func(func)
      , _hasKeywordArgs(hasKeywordArgs)
      , isLocalFunc(isLocal)
      , isRtdAsync(false)
      , isAsync(false)
    {
      _info->name = name.empty() 
        ? py::wstr(func.attr("__name__"))
        : name;

      _info->help = help;
      _info->category = category;
      
      if (!func.ptr() || func.is_none())
        XLO_THROW(L"No python function specified for {0}", name);

      setFuncType(*this, features, isVolatile);

      if (isAsync)
      {
        _info->args.resize(numArgs + 1);
        _info->args[0] = FuncArg(nullptr, nullptr, FuncArg::AsyncHandle);
        for (auto i = 1u; i < numArgs + 1; ++i)
          _args.push_back(PyFuncArg(_info, i));
      }
      else
      {
        _info->args.resize(numArgs);
        for (auto i = 0u; i < numArgs; ++i)
          _args.push_back(PyFuncArg(_info, i));
      }

      _numPositionalArgs = (uint16_t)(_args.size() - (_hasKeywordArgs ? 1u : 0));
    }

    PyFuncInfo::~PyFuncInfo()
    {
      py::gil_scoped_acquire getGil;
      returnConverter.reset();
      _args.clear();
      _func = py::object();
    }

    void PyFuncInfo::setFuncOptions(unsigned val)
    {
      _info->options = val;
    }

    void PyFuncInfo::setReturnConverter(const std::shared_ptr<const IPyToExcel>& conv)
    {
      returnConverter = conv;
    }

    void PyFuncInfo::convertArgs(
      const ExcelObj** xlArgs, 
      PyObject** args,
      py::object& kwargs) const
    {
      args = args + theVectorCallOffset;
      uint16_t i = 0;
      try
      {
        for (; i < _numPositionalArgs; ++i)
        {
          auto* defaultValue = _args[i].getDefault().ptr();
          auto* pyObj = (*_args[i].converter)(*xlArgs[i], defaultValue);
          args[i] = pyObj;
        }

        if (_hasKeywordArgs)
          kwargs = PySteal<py::object>(readKeywordArgs(*xlArgs[_numPositionalArgs]));
      }
      catch (const std::exception& e)
      {
        // Unwind any args already written
        for (auto j = 0u; j < i; ++j)
          Py_DECREF(args[j]);
        
        // We give the arg number 1-based as it's more natural
        XLO_THROW(L"Error in arg {1} '{0}': {2}",
          _args[i].arg.name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
      }
    }

    py::object PyFuncInfo::invoke(PyObject* const* args, const size_t nArgs, PyObject* kwargs) const
    {
#if PY_VERSION_HEX < 0x03080000
      auto argTuple = PySteal<py::tuple>(PyTuple_New(nArgs));
      for (auto i = 0u; i < nArgs; ++i)
        PyTuple_SET_ITEM(argTuple.ptr(), i, args[i]);

      auto retVal = _hasKeywordArgs
        ? PyObject_Call(_func.ptr(), argTuple.ptr(), kwargs)
        : PyObject_CallObject(_func.ptr(), argTuple.ptr());
#else
      auto retVal = _PyObject_FastCallDict(
        _func.ptr(), args + theVectorCallOffset, nArgs | PY_VECTORCALL_ARGUMENTS_OFFSET, kwargs);
#endif
      return PySteal<>(retVal);
    }

    void PyFuncInfo::invoke(
      ExcelObj& result, 
      PyObject* const* args,
      PyObject* kwargsDict) const noexcept
    {
      try
      {
        assert(!!kwargsDict == _hasKeywordArgs);

        auto retVal = invoke(args, _numPositionalArgs, kwargsDict);

        result = returnConverter
          ? (*returnConverter)(*retVal.ptr())
          : FromPyObj()(retVal.ptr());
      }
      catch (const py::error_already_set& e)
      {
        raiseUserException(e);
        result = e.what();
      }
      catch (const std::exception& e)
      {
        result = e.what();
      }
    }

    template<bool TThreadSafe>
    ExcelObj* pythonCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        py::gil_scoped_acquire gilAcquired;
        PyErr_Clear();

        std::array<PyObject*, PyFuncInfo::theVectorCallOffset + XL_MAX_UDF_ARGS> argsArray;
        py::object kwargs;

        info->convertArgs(xlArgs, argsArray, kwargs);

        // This finally block seems pretty heavy, but an array<py::object> would result in 
        // 256 dtor calls every invocation. This approach does just what is required.
        auto finally = [
          begin = argsArray.begin() + PyFuncInfo::theVectorCallOffset, 
          end = argsArray.begin() + info->argArraySize()
        ](void*)
        {
          for (auto i = begin; i != end; ++i)
            Py_DECREF(*i);
        };
        unique_ptr<void, decltype(finally)> cleanup(0, finally);

        if constexpr (TThreadSafe)
        {
          ExcelObj result;
          info->invoke(result, argsArray, kwargs.ptr());
          return returnValue(std::move(result));
        }
        else
        {
          // Static OK since we have the GIL and are single-threaded so can
          // only be called on Excel's main thread. The args array is large 
          // enough: registration with Excel will fail otherwise.
          static ExcelObj result;
          info->invoke(result, argsArray, kwargs.ptr());
          return &result;
        }
      }
      catch (const py::error_already_set& e)
      {
        raiseUserException(e);
        return returnValue(e.what());
      }
      catch (const std::exception& e)
      {
        return returnValue(e.what());
      }
      catch (...)
      {
        return returnValue(CellError::Null);
      }
    }

    shared_ptr<const DynamicSpec> PyFuncInfo::createSpec(const std::shared_ptr<const PyFuncInfo>& func)
    {
      // We implement this as a static function taking a shared_ptr rather than using 
      // shared_from_this with PyFuncInfo as the latter causes pybind to catch a 
      // std::bad_weak_ptr during construction which seems rather un-C++ like and irksome
      func->checkArgConverters();
      if (func->isAsync)
        return make_shared<DynamicSpec>(func->info(), &pythonAsyncCallback, func);
      else if (func->isRtdAsync)
        return make_shared<DynamicSpec>(func->info(), &pythonRtdCallback, func);
      else if (func->isThreadSafe())
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<true>, func);
      else
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<false>, func);
    }

    void PyFuncInfo::checkArgConverters() const
    {
      // TODO: handle kwargs#
      for (auto i = 0u; i < _args.size() - (_hasKeywordArgs ? 1u : 0); ++i)
        if (!_args[i].converter)
          XLO_THROW(L"Converter not set in func '{}' for arg '{}'", info()->name, _args[i].getName());
    }

   

    //TODO: Refactor Python FileSource
    // It might be better for lifetime management if the whole FileSource interface was exposed
    // via the core, then a reference to the FileSource can be held and closed by the module itself
    class RegisteredModule : public WatchedSource
    {
    public:
      /// <summary>
      /// If provided, a linked workbook can be used for local functions
      /// </summary>
      /// <param name="modulePath"></param>
      /// <param name="workbookName"></param>
      RegisteredModule(
        const wstring& modulePath,
        const wchar_t* workbookName)
        : WatchedSource(
            modulePath.empty() ? XLOPY_ANON_SOURCE : modulePath.c_str(),
            workbookName)
      {
        _linkedWorkbook = workbookName;
      }

      ~RegisteredModule()
      {
        try
        {
          if (!_module)
            return;

          // TODO: cancel running async tasks?
          py::gil_scoped_acquire getGil;

          // Call module cleanup function
          if (py::hasattr(_module, XLOPY_CLEANUP_FUNCTION))
            _module.attr(XLOPY_CLEANUP_FUNCTION)();
         
          auto success = unloadModule(_module.release());

          XLO_DEBUG(L"Python module unload {1} for '{0}'", 
            sourceName(), success ? L"succeeded" : L"failed");
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Error unloading python module '{0}': {1}", 
            utf16ToUtf8(sourceName()), e.what());
        }
      }

      void registerPyFuncs(
        const py::handle& pyModule,
        const vector<shared_ptr<PyFuncInfo>>& functions,
        const bool append)
      {
        // This function takes a handle from .release() rather than a py::object
        // to avoid needing the GIL to change the refcount.
        _module = py::reinterpret_steal<py::object>(pyModule);
        vector<shared_ptr<const WorksheetFuncSpec>> nonLocal, localFuncs;

        for (auto& f : functions)
        {
          if (!_linkedWorkbook)
            f->isLocalFunc = false;
          auto spec = PyFuncInfo::createSpec(f);

          if (f->isLocalFunc)
            localFuncs.emplace_back(std::move(spec));
          else
            nonLocal.emplace_back(std::move(spec));
        }

        registerFuncs(nonLocal, append);
        if (!localFuncs.empty())
          registerLocal(localFuncs, append);
      }

      void reload() override
      {
        auto[source, addin] = Python::findSource(sourcePath().c_str());
        if (source.get() != this)
          XLO_THROW(L"Error reloading '{0}': source ptr mismatch", sourcePath());
        
        // Rescan the module, passing in the module handle if it exists
        py::gil_scoped_acquire get_gil;
        if (_module && !_module.is_none())
          addin->thread->callback("xloil.importer", "_import_scan", _module);
        else
          addin->thread->callback("xloil.importer", "_import_file", sourcePath(), addin->pathName(), linkedWorkbook());
      }

    private:
      bool _linkedWorkbook;
      py::object _module;
    };

    std::shared_ptr<RegisteredModule>
      FunctionRegistry::addModule(
        AddinContext& context,
        const std::wstring& modulePath,
        const wchar_t* workbookName)
    {
      auto[source, addin] = FileSource::findSource(modulePath.c_str());
      if (source)
        return std::static_pointer_cast<RegisteredModule>(source);

      auto fileSrc = make_shared<RegisteredModule>(modulePath, workbookName);
      context.addSource(fileSrc);
      return fileSrc;
    }

    namespace
    {
      auto getModulePath(const py::object& module)
      {
        return !module.is_none() && py::hasattr(module, "__file__")
          ? module.attr("__file__").cast<wstring>()
          : L"";
      }
    }

    void registerFunctions(
      const vector<shared_ptr<PyFuncInfo>>& functions,
      py::object& module,
      const py::object& addinCtx,
      const bool append)
    {
      // Called from python so we have the GIL
      // A "null" module handle is used by jupyter
      const auto modulePath = getModulePath(module);

      py::gil_scoped_release releaseGil;

      auto registeredMod = FunctionRegistry::addModule(
        addinCtx.is_none() 
          ? theCoreAddin().context 
          : findAddin(pyToWStr(addinCtx).c_str()).context,
        modulePath, 
        nullptr);
      registeredMod->registerPyFuncs(module.release(), functions, append);
    }

    void deregisterFunctions(
      const py::object& moduleHandle,
      const py::object& functionNames)
    {
      // Called from python so we have the GIL

      const auto modulePath = getModulePath(moduleHandle);

      auto[foundSource, foundAddin] = FileSource::findSource(
        modulePath.empty() ? XLOPY_ANON_SOURCE : modulePath.c_str());

      if (!foundSource)
      {
        XLO_WARN(L"Call to deregisterFunctions with unknown source '{0}'", modulePath);
        return;
      }
      vector<wstring> funcNames;
      auto iter = py::iter(functionNames);
      while (iter != py::iterator::sentinel())
      {
        funcNames.push_back(iter->cast<wstring>());
        ++iter;
      }

      py::gil_scoped_release releaseGil;

      for (auto& func : funcNames)
        foundSource->deregister(func);
    }

    namespace
    {
      void bitSet(int& x, int mask, bool val)
      {
        if (val)
          x |= mask;
        else
          x &= ~mask;
      }

      string pyFuncInfoToString(const PyFuncInfo& info)
      {
        string result = utf16ToUtf8(info.info()->name) + "(";
        for (auto& arg : info.constArgs())
          result += utf16ToUtf8(arg.getName()) + (arg.converter ? formatStr(": %s, ", arg.converter->name()) : ", ");
        if (!info.constArgs().empty())
          result.resize(result.size() - 2);
        result.push_back(')');
        if (info.getReturnConverter())
          result += formatStr(" -> ", typeid(*info.getReturnConverter()).name());
        return result;
      }
      
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyFuncArg>(mod, "FuncArg")
          .def_property("name", &PyFuncArg::getName, &PyFuncArg::setName)
          .def_property("help", &PyFuncArg::getHelp, &PyFuncArg::setHelp)
          .def_readwrite("converter", &PyFuncArg::converter)
          .def_property("default", &PyFuncArg::getDefault, &PyFuncArg::setDefault)
          .def_property("allow_range",
            [](PyFuncArg& x) { return (x.arg.type & FuncArg::Range) != 0; },
            [](PyFuncArg& x, bool v)
            {
              bitSet(x.arg.type, FuncArg::Range, v);
            }
        );

        py::class_<PyFuncInfo, shared_ptr<PyFuncInfo>>(mod, "FuncSpec")
          .def(py::init<py::function, wstring, unsigned, string, wstring, wstring, bool, bool, bool>(),
            py::arg("func"),
            py::arg("name") = "",
            py::arg("nargs"),
            py::arg("features") = py::none(),
            py::arg("help") = "",
            py::arg("category") = "",
            py::arg("local") = true,
            py::arg("volatile") = false,
            py::arg("has_kwargs") = false)
          .def_property("return_converter", &PyFuncInfo::getReturnConverter, &PyFuncInfo::setReturnConverter)
          .def_property_readonly("args", &PyFuncInfo::args)
          .def_property_readonly("name", [](const PyFuncInfo& self) { return self.info()->name; })
          .def_property_readonly("help", [](const PyFuncInfo& self) { return self.info()->help; })
          .def("__str__", pyFuncInfoToString);

        mod.def("register_functions", &registerFunctions, 
          py::arg("funcs"),
          py::arg("module")=py::none(),
          py::arg("addin")=py::none(),
          py::arg("append")=false);
        mod.def("deregister_functions", &deregisterFunctions);
      });
    }
  }
}