#include "FunctionRegister.h"
#include "InjectedModule.h"
#include "Main.h"
#include "BasicTypes.h"
#include "Dictionary.h"
#include <xloil/ExcelCall.h>
#include <xloil/AsyncHelper.h>
#include <xloil/ExcelState.h>
#include <pybind11/stl.h>
#include <CTPL/ctpl_stl.h>
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
namespace py = pybind11;
using namespace pybind11::literals;

namespace xloil
{
  namespace Python
  {
    struct PyFuncInfo
    {
      PyFuncInfo(const shared_ptr<FuncInfo>& info, const py::function& func, bool hasKeywordArgs)
      {
        this->info = info;
        this->func = func;
        this->hasKeywordArgs = hasKeywordArgs;
        this->isLocalFunc = false;
        argConverters.resize(info->numArgs() - (hasKeywordArgs ? 1 : 0));
      }

      void setArgTypeDefault(size_t i, shared_ptr<IPyFromExcel> converter, py::object defaultVal)
      {
        argConverters[i] = std::make_pair(converter, defaultVal);
      }

      void setArgType(size_t i, shared_ptr<IPyFromExcel> converter)
      {
        argConverters[i] = std::make_pair(converter, py::object());
      }

      shared_ptr<FuncInfo> info;
      py::function func;
      bool hasKeywordArgs;
      vector<pair<shared_ptr<IPyFromExcel>, py::object>> argConverters;
      shared_ptr<IPyToExcel> returnConverter;
      bool isLocalFunc;

      void setFuncOptions(int val)
      {
        info->options = val;
      }
    };


    pair<py::tuple, py::object> convertArgsToPython(PyFuncInfo* info, const ExcelObj** xlArgs)
    {
      auto nArgs = info->argConverters.size();
      auto pyArgs = PySteal<py::tuple>(PyTuple_New(nArgs));

      // TODO: is it worth having a enum switch to convert primitive types rather than a v-call
      for (auto i = 0; i < nArgs; ++i)
      {
        try
        {
          auto* def = info->argConverters[i].second.ptr();
          auto* pyObj = (*info->argConverters[i].first)(*xlArgs[i], def);
          PyTuple_SET_ITEM(pyArgs.ptr(), i, pyObj);
        }
        catch (const std::exception& e)
        {
          // TODO: could we explain what type is required?
          // We give the arg number 1-based as it's more natural
          XLO_THROW(L"Error reading '{0}' arg #{1}: {2}",
            info->info->args[i].name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
        }
      }
      if (info->hasKeywordArgs)
      {
        auto kwargs = PySteal<py::dict>(readKeywordArgs(*xlArgs[nArgs]));
        return make_pair(pyArgs, kwargs);
      }
      else
        return make_pair(pyArgs, py::none());
    }

    ExcelObj* invokePyFunction(PyFuncInfo* info, PyObject* args, PyObject* kwargs)
    {
      try
      {
        py::object ret;
        if (kwargs != Py_None)
          ret = PySteal<py::object>(PyObject_Call(info->func.ptr(), args, kwargs));
        else
          ret = PySteal<py::object>(PyObject_CallObject(info->func.ptr(), args));

        // TODO: Review this if we ever go to multi-threaded python
        static ExcelObj result;

        result = info->returnConverter
          ? (*info->returnConverter)(*ret.ptr())
          : FromPyObj()(ret.ptr());

        return &result;
      }
      catch (const std::exception& e)
      {
        return ExcelObj::returnValue(e.what());
      }
    }

    ExcelObj* pythonCallback(PyFuncInfo* info, const ExcelObj** xlArgs)
    {
      try
      {
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        auto[args, kwargs] = convertArgsToPython(info, xlArgs);

        return invokePyFunction(info, args.ptr(), kwargs.ptr());
      }
      catch (const std::exception& e)
      {
        return ExcelObj::returnValue(e.what());
      }
      catch (...)
      {
        return ExcelObj::returnValue("#ERROR");
      }
    }

    struct ThreadContext
    {
      ThreadContext() 
        : _startTime(GetTickCount64())
      {}
      bool cancelled()
      {
        return _startTime < lastCalcCancelledTicks() || yieldAndCheckIfEscPressed();
      }
    private:
      size_t _startTime;
    };

    static ctpl::thread_pool* thePythonWorkerThread = nullptr;

    void pythonAsyncCallback(PyFuncInfo* info, const ExcelObj* asyncHandle, const ExcelObj** xlArgs)
    {
      try
      {
        py::gil_scoped_acquire gilAcquired;
        {
          PyErr_Clear();

          // I think it's better to process the arguments to python here rather than 
          // copying the ExcelObj's and converting on the async thread (since CPython
          // is single threaded anyway)
          auto[args, kwargs] = convertArgsToPython(info, xlArgs);
          if (kwargs.is_none())
            kwargs = py::dict();

          kwargs["xloil_thread_context"] = ThreadContext();

          // Need to drop pybind links before capturing in lambda otherwise the destructor
          // is called at some random time after losing the GIL and it crashes.
          auto argsP = args.release().ptr();
          auto kwargsP = kwargs.release().ptr();
         
          auto functor = AsyncHolder(
            [info, argsP, kwargsP]() mutable
            {
              py::gil_scoped_acquire gilAcquired;
              {
                auto ret = invokePyFunction(info, argsP, kwargsP);
                Py_XDECREF(argsP);
                Py_XDECREF(kwargsP);
                return ret;
              }
            },
            asyncHandle);
          thePythonWorkerThread->push(functor);
        }
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(e.what());
      }
      catch (...)
      {
        XLO_ERROR("Async unknow error");
      }
    }

    void registerFunc(const shared_ptr<PyFuncInfo>& funcInfo)
    {
      if (funcInfo->info->options & FuncInfo::ASYNC)
      {
        if (!thePythonWorkerThread)
          thePythonWorkerThread = new ctpl::thread_pool(1);

        theCore->registerFunc(funcInfo->info, &pythonAsyncCallback, funcInfo);
      }
      else
        theCore->registerFunc(funcInfo->info, &pythonCallback, funcInfo);
    }

    void handleFileChange(const wchar_t* dirName, const wchar_t* fileName, const FileAction action);

    class RegisteredModule
    {
    public:
      RegisteredModule(const wstring& modulePath, const wchar_t* workbookModule)
        : _modulePath(modulePath)
      {
        const auto path = fs::path(modulePath);
        _fileWatcher = std::static_pointer_cast<const void>
          (Event_DirectoryChange(fs::path(path).remove_filename()).bind(handleFileChange));
        if (workbookModule)
          _workbookModule = workbookModule;
      }
      ~RegisteredModule()
      {
        XLO_DEBUG(L"Deregistering functions in module '{0}'", _modulePath);
        for (auto& f : _functions)
          theCore->deregister(f.second->info->name);
        if (!_workbookModule.empty())
          theCore->forgetLocal(_workbookModule.c_str());
      }

      void registerFuncs(const vector<shared_ptr<PyFuncInfo>>& functions)
      {
        if (_functions.empty())
        {
          // Fresh registration, just add functions
          for (auto& f : functions)
          {
            if (f->isLocalFunc) continue;
            _functions.emplace(f->info->name, f);
            registerFunc(f);
          }
        }
        else
        {
          // Trickier case: potentially re-registering functions
          map<wstring, shared_ptr<PyFuncInfo>> newMap;

          for (auto& f : functions)
          {
            if (f->isLocalFunc) continue;
            auto iFunc = _functions.find(f->info->name);

            // If the function name already exists, try to avoid re-registering
            if (iFunc != _functions.end())
            {
              // Attempt to patch the function context to refer to to the new py function
              if (!theCore->reregister(iFunc->second->info, std::static_pointer_cast<void>(f)))
              {
                // If that failed, we need to do it ourselves
                registerFunc(f);
              }
              // Having handled this function, remove it from the old map
              _functions.erase(iFunc);
            }
            else
              registerFunc(f);
            newMap.emplace(f->info->name, f);
          }

          // Any functions remaining in the old map must have been removed from the module
          // so we can deregister them, but if that fails we have to keep them or they
          // will be orphaned
          for (auto& f : _functions)
            if (!theCore->deregister(f.second->info->name))
              newMap.emplace(f);

          _functions = newMap;

          if (!_workbookModule.empty())
            registerLocalFuncs(_workbookModule.c_str(), functions);
        }
      }

      void registerLocalFuncs(
        const wchar_t* workbookName,
        const vector<shared_ptr<PyFuncInfo>>& functions)
      {
        vector<shared_ptr<const FuncInfo>> funcInfo;
        vector<ExcelFuncPrototype> funcs;
        for (auto &f : functions)
        {
          if (f->isLocalFunc)
          {
            funcInfo.push_back(f->info);
            funcs.push_back([f](const FuncInfo&, const ExcelObj** args)
            {
              return pythonCallback(f.get(), args);
            });
          }
        }
        theCore->registerLocal(workbookName, funcInfo, funcs);
      }

      const wstring& modulePath() const { return _modulePath; }
      bool workbookModule() const { return !_workbookModule.empty(); }

    private:
      map<wstring, shared_ptr<PyFuncInfo>> _functions;
      shared_ptr<const void> _fileWatcher;
      wstring _modulePath;
      wstring _workbookModule;
    };

    FunctionRegistry& FunctionRegistry::get() 
    {
      static FunctionRegistry instance;
      return instance;
    }


    std::shared_ptr<RegisteredModule> 
      FunctionRegistry::addModule(const std::wstring& modulePath, const wchar_t* workbookName)
    {
      auto it = _modules.find(modulePath);
      if (it == _modules.end())
       it = _modules.insert(make_pair(modulePath, make_shared<RegisteredModule>(modulePath, workbookName))).first;
      return it->second;
    }

    std::shared_ptr<RegisteredModule> 
      FunctionRegistry::addModule(const pybind11::module& moduleHandle, const wchar_t* workbookName)
    {
      return addModule(moduleHandle.attr("__file__").cast<wstring>(), workbookName);
    }

    FunctionRegistry::FunctionRegistry()
    {
      static auto handler = Event_PyBye().bind([] 
      {
        FunctionRegistry::get().modules().clear(); 
        if (thePythonWorkerThread)
          delete thePythonWorkerThread;
      });
    }

    void handleFileChange(const wchar_t* dirName, const wchar_t* fileName, const FileAction action)
    {
      auto filePath = (fs::path(dirName) / fileName).wstring();
      auto& registry = FunctionRegistry::get().modules();
      auto found = registry.find(filePath);
      if (found == registry.end())
        return;
      switch (action)
      {
      case FileAction::Modified:
        XLO_INFO(L"Module '{0}' modified, reloading.", filePath);
        scanModule(py::wstr(filePath));
        break;
      case FileAction::Delete:
        XLO_INFO(L"Module '{0}' deleted, removing functions.", filePath);
        registry.erase(filePath);
        break;
      }
    }

    void registerFunctions(
      const py::object& moduleHandle, 
      const vector<shared_ptr<PyFuncInfo>>& functions)
    {
      FunctionRegistry::get().addModule(moduleHandle.cast<py::module>())
        ->registerFuncs(functions);
    }

    void writeToLog(const char* message, const char* level)
    {
      SPDLOG_LOGGER_CALL(spdlog::default_logger_raw(), spdlog::level::from_str(level), message);
    }
    
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<FuncArg>(mod, "FuncArg")
          .def(py::init<const wchar_t*, const wchar_t*>())
          .def_readwrite("name", &FuncArg::name)
          .def_readwrite("help", &FuncArg::help)
          .def_readwrite("allow_range", &FuncArg::allowRange);

        py::class_<FuncInfo, shared_ptr<FuncInfo>>(mod, "FuncInfo")
          .def(py::init())
          .def_readwrite("name", &FuncInfo::name)
          .def_readwrite("help", &FuncInfo::help)
          .def_readwrite("category", &FuncInfo::category)
          .def_readwrite("args", &FuncInfo::args);

        py::enum_<FuncInfo::FuncOpts>(mod, "FuncOpts", py::arithmetic())
          .value("Async", FuncInfo::ASYNC)
          .value("Macro", FuncInfo::MACRO_TYPE)
          .value("ThreadSafe", FuncInfo::THREAD_SAFE)
          .value("Volatile", FuncInfo::VOLATILE)
          .export_values();

        py::class_<PyFuncInfo, shared_ptr<PyFuncInfo>>(mod, "FuncHolder")
          .def(py::init<const shared_ptr<FuncInfo>&, const py::function&, bool>())
          .def("set_arg_type", &PyFuncInfo::setArgType, py::arg("i"), py::arg("arg_type"))
          .def("set_arg_type_defaulted", &PyFuncInfo::setArgTypeDefault, py::arg("i"), py::arg("arg_type"), py::arg("default"))
          .def("set_opts", &PyFuncInfo::setFuncOptions, py::arg("flags"))
          .def_readwrite("local", &PyFuncInfo::isLocalFunc);

        py::class_<ThreadContext>(mod, "ThreadContext")
          .def("cancelled", &ThreadContext::cancelled);
       

        mod.def("in_wizard", &xloil::inFunctionWizard);
        mod.def("register_functions", &registerFunctions);
        mod.def("log", &writeToLog, py::arg("msg"), py::arg("level")="info");
      });
    }
  }
}