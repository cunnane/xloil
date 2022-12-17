#include "PyFunctionRegister.h"
#include "PyCore.h"
#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "TypeConversion/PyDictType.h"
#include "PySource.h"
#include "AsyncFunctions.h"
#include "PyEvents.h"
#include <xloil/StaticRegister.h>
#include <xloil/DynamicRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Caller.h>
#include <xloil/FPArray.h>
#include <xloil/RtdServer.h>
#include <xlOil/ExcelThread.h>
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
using std::wstring_view;
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

    unsigned readFuncFeatures(
      const string& features,
      PyFuncInfo& info,
      bool isVolatile,
      bool& isLocalFunc)
    {
      unsigned funcOpts = isVolatile ? FuncInfo::VOLATILE : 0;

      if (features.find("macro") != string::npos)
        funcOpts |= FuncInfo::MACRO_TYPE;
      if (features.find("fastarray") != string::npos)
        funcOpts |= FuncInfo::ARRAY;
      if (features.find("command") != string::npos)
        funcOpts |= FuncInfo::COMMAND;
      if (features.find("threaded") != string::npos)
      {
        funcOpts |= FuncInfo::THREAD_SAFE;
        isLocalFunc = false;
      }
      if (features.find("rtd") != string::npos)
      {
        info.isRtdAsync = true;
      }
      if (features.find("async") != string::npos)
      {
        info.isAsync = true;
        isLocalFunc = false;
        if (funcOpts > 0)
          XLO_THROW("Async cannot be used with other function features like command, macro, etc");
      }

      return funcOpts;
    }

    PyFuncInfo::PyFuncInfo(
      const pybind11::function& func,
      const std::vector<PyFuncArg>& args,
      const std::wstring& name,
      const std::string& features,
      const std::wstring& help,
      const std::wstring& category,
      bool isLocal,
      bool isVolatile)
      : _info(new FuncInfo())
      , _func(func)
      , _args(args)
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

      py::gil_scoped_release releaseGil;

      _info->options = readFuncFeatures(features, *this, isVolatile, isLocalFunc);

      if (!_info->isValid())
        XLO_THROW("Invalid combination of function features: '{}'", features);
    }

    PyFuncInfo::~PyFuncInfo()
    {
      py::gil_scoped_acquire getGil;
      returnConverter.reset();
      _args.clear();
      _func = py::object();
    }

    void PyFuncInfo::setReturnConverter(const std::shared_ptr<const IPyToExcel>& conv)
    {
      returnConverter = conv;
    }


    struct CommandReturn
    {
      using return_type = int;

      CommandReturn(const IPyToExcel*) {}

      int operator()(PyObject* /*retVal*/) const
      {
        return 1; // Ignore return value
      }

      int operator()(const char* err, const PyFuncInfo* info) const
      {
        XLO_ERROR(L"{0}: {1}", info->name(), utf8ToUtf16(err));
        return 0;
      }

      int operator()(CellError, const PyFuncInfo* info) const
      {
        XLO_ERROR(L"{0}: unknown error", info->name());
        return 0;
      }
    };

    struct FPArrayReturn
    {
      using return_type = FPArray*;

      FPArrayReturn(const IPyToExcel*) {}

      FPArray* operator()(PyObject* retVal) const
      {
        return numpyToFPArray(*retVal).get();
      }

      FPArray* operator()(const char* err, const PyFuncInfo* info) const
      {
        XLO_ERROR(L"{0}: {1}", info->name(), utf8ToUtf16(err));
        return FPArray::empty();
      }

      FPArray* operator()(CellError, const PyFuncInfo* info) const
      {
        XLO_ERROR(L"{0}: unknown error", info->name());
        return FPArray::empty();
      }
    };

    struct ExcelObjReturn
    {
      using return_type = ExcelObj*;

      const IPyToExcel* _returnConverter;

      ExcelObjReturn(const IPyToExcel* returnConverter)
        : _returnConverter(returnConverter)
      {}

      ExcelObj* operator()(PyObject* retVal) const
      {
        // Static is OK since we have the GIL and are single-threaded so can
        // must be on Excel's main thread. The args array is large enough:
        // we cannot register a function with more arguments than that.
        static ExcelObj result;

        result = _returnConverter
          ? (*_returnConverter)(*retVal)
          : FromPyObj()(retVal);

        return (*this)(result);
      }

      template<class T> ExcelObj* operator()(T&& x, const PyFuncInfo* = nullptr) const
      {
        // No need to include function name in error messages since the calling function
        // is usually clear in a worksheet
        return returnValue(std::forward<T>(x));
      }
    };

    struct ExcelObjThreadSafeReturn
    {
      using return_type = ExcelObj*;

      const IPyToExcel* _returnConverter;

      ExcelObjThreadSafeReturn(const IPyToExcel* returnConverter)
        : _returnConverter(returnConverter)
      {}

      ExcelObj* operator()(PyObject* retVal) const
      {
        return returnValue(_returnConverter
          ? (*_returnConverter)(*retVal)
          : FromPyObj()(retVal));
      }

      template<class T> ExcelObj* operator()(T&& x, const PyFuncInfo* = nullptr) const
      {
        // No need to include function name in error messages since the calling function
        // is usually clear in a worksheet
        return returnValue(std::forward<T>(x));
      }
    };

    template<class TReturn = ExcelObjReturn>
    typename TReturn::return_type pythonCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      TReturn returner(info->getReturnConverter().get());

      try
      {
        py::gil_scoped_acquire gilAcquired;
        PyErr_Clear(); // TODO: required?
        return returner(info->invoke([&](auto i) -> auto& { return *xlArgs[i]; }).ptr());
      }
      catch (const py::error_already_set& e)
      {
        raiseUserException(e);
        return returner(e.what(), info);
      }
      catch (const std::exception& e)
      {
        return returner(e.what(), info);
      }
      catch (...)
      {
        return returner(CellError::Value, info);
      }
    }

    shared_ptr<const DynamicSpec>
      PyFuncInfo::createSpec(
        const std::shared_ptr<PyFuncInfo>& func)
    {
      // We implement this as a static function taking a shared_ptr rather than using 
      // shared_from_this with PyFuncInfo as the latter causes pybind to catch a 
      // std::bad_weak_ptr during construction which seems rather un-C++ like and irksome
      func->writeExcelArgumentDescription();
      auto cfunc = std::const_pointer_cast<const PyFuncInfo>(func);
      if (func->isAsync)
        return make_shared<DynamicSpec>(func->info(), &pythonAsyncCallback, cfunc);
      else if (func->isRtdAsync)
        return make_shared<DynamicSpec>(func->info(), &pythonRtdCallback, cfunc);
      else if (func->isThreadSafe())
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<ExcelObjThreadSafeReturn>, cfunc);
      else if (func->isCommand())
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<CommandReturn>, cfunc);
      else if (func->isFPArray())
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<FPArrayReturn>, cfunc);
      else
        return make_shared<DynamicSpec>(func->info(), &pythonCallback<>, cfunc);
    }

    void PyFuncInfo::writeExcelArgumentDescription()
    {
      const auto numArgs = _args.size();
      _numPositionalArgs = (uint16_t)numArgs;

      if (numArgs == 0)
        return;

      auto lastArg = _args.end() - 1;

      if (lastArg->isKeywords())
      {
        _hasKeywordArgs = true;
        --_numPositionalArgs;
        if (_numPositionalArgs > 0)
          --lastArg;
      }

      if (lastArg->isVargs())
      {
        if (_hasKeywordArgs)
          std::iter_swap(lastArg, _args.end() - 1);
        _hasVariableArgs = true;
        --_numPositionalArgs;
      }

      auto& registerArgs = _info->args;

      registerArgs.reserve(
        numArgs
        + (isAsync ? 1 : 0) 
        + (_hasVariableArgs ? 100 : 0));

      if (isAsync)
        registerArgs.emplace_back(wstring_view(), wstring_view(), FuncArg::AsyncHandle);

      for (auto& arg : _args)
      {
        int flags = FuncArg::Obj;

        if (arg.type.find("array") != string::npos)
          flags = FuncArg::Array;
        else if (arg.type.find("range") != string::npos)
          flags |= FuncArg::Range;

        if (arg.default)
          flags |= FuncArg::Optional;

        if ((flags & FuncArg::Array) != 0)
          arg.converter.reset(createFPArrayConverter());

        // If no help string has been provided, give a type hint based on 
        // the arg converter
        if (arg.help.empty() && arg.converter)
        {
          wstring argType = utf8ToUtf16(arg.converter->name());
          if (arg.default)
          {
            arg.help = formatStr(L"[%s]", argType.c_str());
            // Could do this but we'd need to reacquire the GIL
            //arg.help = formatStr(L"[%s]='%s'", argType.c_str(), pyToWStr(arg.default).c_str());
          }
          else
            arg.help = formatStr(L"<%s>", argType.c_str());
        }

        if (!arg.converter && !arg.isKeywords())
          XLO_THROW(L"Converter not set in func '{}' for arg '{}'", info()->name, arg.name);

        registerArgs.emplace_back(
          arg.name,
          arg.help,
          flags);
      }

      if (_hasVariableArgs)
      {
        // Note that having 255 args means the concatenated argument names will certainly
        // exceed 255 chars, which will generate a notification in the log file
        auto numVariableArgs = (isLocalFunc ? XL_MAX_VBA_FUNCTION_ARGS : XL_MAX_UDF_ARGS) - numArgs;
        auto varArgType = registerArgs.back().type | FuncArg::Optional;
        for (size_t i = 1; i < numVariableArgs; ++i)
        {
          registerArgs.emplace_back(
            isLocalFunc ? formatStr(L"a%d", i) : L"",
            formatStr(L"[%s-%d]", _args.back().name, i),
            varArgType);
        }
      }
    }

    //TODO: Refactor Python FileSource
    // It might be better for lifetime management if the whole FileSource interface was exposed
    // via the core, then a reference to the FileSource can be held and closed by the module itself

    /// <summary>
    /// If provided, a linked workbook can be used for local functions
    /// </summary>
    /// <param name="modulePath"></param>
    /// <param name="workbookName"></param>
    RegisteredModule::RegisteredModule(
      const wstring& modulePath,
      const wchar_t* workbookName)
      : LinkedSource(
        modulePath.empty() ? XLOPY_ANON_SOURCE : modulePath.c_str(),
        true,
        workbookName)
    {
      _linkedWorkbook = workbookName;
    }

    RegisteredModule::~RegisteredModule()
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
          filename(), success ? L"succeeded" : L"failed");
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Error unloading python module '{0}': {1}",
          utf16ToUtf8(filename()), e.what());
      }
    }

    void RegisteredModule::registerPyFuncs(
      const py::handle& pyModule,
      const vector<shared_ptr<PyFuncInfo>>& functions,
      const bool append)
    {
      // This function takes a handle from .release() rather than a py::object
      // to avoid needing the GIL to change the refcount.
      _module = py::reinterpret_steal<py::object>(pyModule);
      vector<shared_ptr<const WorksheetFuncSpec>> nonLocal, localFuncs;

      bool usesRtdAsync = false;

      for (auto& f : functions)
      {
        if (!_linkedWorkbook)
          f->isLocalFunc = false;
        auto spec = PyFuncInfo::createSpec(f);

        if (f->isLocalFunc)
          localFuncs.emplace_back(std::move(spec));
        else
          nonLocal.emplace_back(std::move(spec));

        if (f->isRtdAsync)
          usesRtdAsync = true;
      }

      // Prime the RTD pump now as a background task to avoid it blocking 
      // in calculation later.
      if (usesRtdAsync)
        runExcelThread([]() { rtdAsync(shared_ptr<IRtdAsyncTask>()); });

      registerFuncs(nonLocal, append);
      if (!localFuncs.empty())
        registerLocal(localFuncs, append);
    }

    void RegisteredModule::reload()
    {
      auto [source, addin] = Python::findSource(name().c_str());
      if (source.get() != this)
        XLO_THROW(L"Error reloading '{0}': source ptr mismatch", name());

      // Rescan the module, passing in the module handle if it exists
      py::gil_scoped_acquire get_gil;
      if (_module && !_module.is_none())
        addin->importModule(_module);
      else
        addin->importFile(name().c_str(), linkedWorkbook().c_str());
    }

    void RegisteredModule::renameWorkbook(const wchar_t* newPathName)
    {
      if (!_linkedWorkbook) // Should never be called without a linked wb
        return;

      // This is all great but...background thread?
      auto [source, addin] = Python::findSource(name().c_str());

      AddinContext::deleteSource(shared_from_this());

      const auto newSourcePath = addin->getLocalModulePath(newPathName);
      const auto& currentSourcePath = name();

      std::error_code ec;
      if (fs::copy_file(currentSourcePath, newSourcePath, ec))
      {
        const auto wbName = wcsrchr(newPathName, '\\') + 1;
        auto newSource = make_shared<RegisteredModule>(newSourcePath, wbName);
        addin->context.addSource(newSource);
        py::gil_scoped_acquire get_gil;
        addin->importFile(newSourcePath.c_str(), newPathName);
      }
      else
        XLO_WARN(L"On workbook rename, failed to copy source '{0}' to '{1}' because: {3}",
          currentSourcePath, newSourcePath, utf8ToUtf16(ec.message()));
    }

    std::shared_ptr<RegisteredModule>
      FunctionRegistry::addModule(
        AddinContext& context,
        const std::wstring& modulePath,
        const wchar_t* workbookName)
    {
      auto [source, addin] = AddinContext::findSource(modulePath.c_str());
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
      const py::object& addinContext,
      const bool append)
    {
      // Called from python so we have the GIL
      // A "null" module handle is used by jupyter
      const auto modulePath = getModulePath(module);

      auto& context = addinContext.is_none()
        ? theCoreAddin()->context
        : py::cast<PyAddin&>(addinContext).context;

      py::gil_scoped_release releaseGil;
      auto registeredMod = FunctionRegistry::addModule(
        context,
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

      auto [foundSource, foundAddin] = AddinContext::findSource(
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
      string pyFuncInfoToString(const PyFuncInfo& info)
      {
        string result = utf16ToUtf8(info.info()->name) + "(";
        for (auto& arg : info.args())
          result += utf16ToUtf8(arg.name) 
            + (arg.converter ? formatStr(": %s, ", arg.converter->name()) : ", ");
        if (!info.args().empty())
          result.resize(result.size() - 2);
        result.push_back(')');
        if (info.getReturnConverter())
          result += formatStr(" -> %s", info.getReturnConverter()->name());
        return result;
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyFuncArg>(mod, "_FuncArg")
          .def(py::init<>())
          .def_readwrite("name", &PyFuncArg::name)
          .def_readwrite("help", &PyFuncArg::help)
          .def_readwrite("converter", &PyFuncArg::converter)
          .def_readwrite("default", &PyFuncArg::default)
          .def_readwrite("special_type", &PyFuncArg::type);

        py::class_<PyFuncInfo, shared_ptr<PyFuncInfo>>(mod, "_FuncSpec")
          .def(py::init<py::function, vector<PyFuncArg>, wstring, string, wstring, wstring, bool, bool>(),
            py::arg("func"),
            py::arg("args"),
            py::arg("name") = "",
            py::arg("features") = py::none(),
            py::arg("help") = "",
            py::arg("category") = "",
            py::arg("local") = true,
            py::arg("volatile") = false)
          .def_property("return_converter",
            &PyFuncInfo::getReturnConverter,
            &PyFuncInfo::setReturnConverter)
          .def_property_readonly("args",
            [](const PyFuncInfo& self) { return self.args(); })
          .def_property_readonly("name",
            [](const PyFuncInfo& self) { return self.info()->name; })
          .def_property_readonly("help",
            [](const PyFuncInfo& self) { return self.info()->help; })
          .def_property("func",
            &PyFuncInfo::func, &PyFuncInfo::setFunc,
            R"(
              Yes you can change the function which is called by Excel! Use
              with caution.
            )")
          .def_property_readonly("is_threaded",
            &PyFuncInfo::isThreadSafe,
            R"(
              True if the function can be multi-threaded during Excel calcs
            )")
          .def_property_readonly("is_rtd",
            [](const PyFuncInfo& self) { return self.isRtdAsync; },
            R"(
              True if the function uses RTD to provide async returns
            )")
          .def_property_readonly("is_async",
            [](const PyFuncInfo& self) { return self.isAsync; },
            R"(
              True if the function used Excel's native async
            )")
          .def("__str__", pyFuncInfoToString);

        mod.def("_register_functions", &registerFunctions,
          py::arg("funcs"),
          py::arg("module") = py::none(),
          py::arg("addin") = py::none(),
          py::arg("append") = false);

        mod.def("deregister_functions",
          &deregisterFunctions,
          R"(
            Deregisters worksheet functions linked to specified module. Generally, there
            is no need to call this directly.
          )");
      }, 20); // Need to declare before PyAddin
    }
  }
}