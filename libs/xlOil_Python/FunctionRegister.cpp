#include "FunctionRegister.h"
#include "InjectedModule.h"
#include "Main.h"
#include "BasicTypes.h"
#include "Dictionary.h"
#include "File.h"
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Register/AsyncHelper.h>
#include <xloil/Caller.h>
#include <xloil/RtdServer.h>
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
      PyFuncInfo(
        const shared_ptr<FuncInfo>& info, 
        const py::function& func, 
        bool keywordArgs)
      {
        this->info = info;
        this->func = func;
        hasKeywordArgs = keywordArgs;
        isLocalFunc = false;
        isRtdAsync = false;
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
      bool isRtdAsync;

      void setFuncOptions(int val)
      {
        info->options = val;
      }
    };


    pair<py::tuple, py::object> convertArgsToPython(const PyFuncInfo* info, const ExcelObj** xlArgs)
    {
      auto nArgs = info->argConverters.size();
      auto pyArgs = PySteal<py::tuple>(PyTuple_New(nArgs));

      // TODO: is it worth having a enum switch to convert primitive types rather than a v-call
      for (auto i = 0u; i < nArgs; ++i)
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

    void invokePyFunction(ExcelObj& result, const PyFuncInfo* info, PyObject* args, PyObject* kwargs)
    {
      try
      {
        py::object ret;
        if (kwargs != Py_None)
          ret = PySteal<py::object>(PyObject_Call(info->func.ptr(), args, kwargs));
        else
          ret = PySteal<py::object>(PyObject_CallObject(info->func.ptr(), args));

        result = info->returnConverter
          ? (*info->returnConverter)(*ret.ptr())
          : FromPyObj()(ret.ptr());
      }
      catch (const std::exception& e)
      {
        result = e.what();
      }
    }

    ExcelObj* pythonCallback(
      PyFuncInfo* info, 
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        auto[args, kwargs] = convertArgsToPython(info, xlArgs);

        static ExcelObj result; // Ok since python is single threaded
        invokePyFunction(result, info, args.ptr(), kwargs.ptr());
        return &result;
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

    constexpr const char* THREAD_CONTEXT_TAG = "xloil_thread_context";

    struct AsyncNotifier
    {
      AsyncNotifier()
        : _startTime(GetTickCount64())
      {}
      bool cancelled()
      {
        return _startTime < lastCalcCancelledTicks() || yieldAndCheckIfEscPressed();
      }
    private:
      decltype(GetTickCount64()) _startTime;
    };
    
    static ctpl::thread_pool* getWorkerThreadPool()
    {
      constexpr size_t nThreads = 1;
      static auto* workerPool = new ctpl::thread_pool(nThreads);
      static auto workerPoolDeleter = Event_PyBye().bind([]
      {
        if (workerPool)
          delete workerPool;
      });
      return workerPool;
    }

    void pythonAsyncCallback(
      PyFuncInfo* info, 
      const ExcelObj* asyncHandle, 
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        PyObject *argsP, *kwargsP;
        {
          py::gil_scoped_acquire gilAcquired;

          PyErr_Clear();

          // I think it's better to process the arguments to python here rather than 
          // copying the ExcelObj's and converting on the async thread (since CPython
          // is single threaded anyway)
          auto[args, kwargs] = convertArgsToPython(info, xlArgs);
          if (kwargs.is_none())
            kwargs = py::dict();

          kwargs[THREAD_CONTEXT_TAG] = AsyncNotifier();

          // Need to drop pybind links before capturing in lambda otherwise the destructor
          // is called at some random time after losing the GIL and it crashes.
          argsP = args.release().ptr();
          kwargsP = kwargs.release().ptr();
        }
        auto functor = AsyncHolder(
          [info, argsP, kwargsP]() mutable
          {
            py::gil_scoped_acquire gilAcquired;
            {
              ExcelObj result;
              invokePyFunction(result, info, argsP, kwargsP);
              Py_XDECREF(argsP);
              Py_XDECREF(kwargsP);
              return returnValue(std::move(result));
            }
          },
          asyncHandle);
        getWorkerThreadPool()->push(functor);
      }
      catch (const std::exception& e)
      {
        XLO_WARN(e.what());
        asyncReturn(*asyncHandle, ExcelObj(e.what()));
      }
      catch (...)
      {
        XLO_WARN("Async unknown error");
        asyncReturn(*asyncHandle, ExcelObj(CellError::Value));
      }
    }

    struct RtdNotifier
    {
      RtdNotifier(IRtdNotify& n) : notifier(n) {}
      bool cancelled()
      {
        return notifier.isCancelled();
      }
      IRtdNotify& notifier;
    };

    /// <summary>
    /// Holder for python target function and its arguments.
    /// Able to compare arguments with another AsyncTask
    /// </summary>
    struct AsyncTask
    {

      /// <summary>
      /// Steals references to PyObjects
      /// </summary>
      AsyncTask(PyFuncInfo* info, PyObject* args, PyObject* kwargs)
        : _info(info)
        , _args(args)
        , _kwargs(kwargs)
      {}

      PyFuncInfo* _info;
      PyObject *_args, *_kwargs;

      ~AsyncTask()
      {
        py::gil_scoped_acquire gilAcquired;
        Py_XDECREF(_args);
        Py_XDECREF(_kwargs);
      }

      std::future<void> operator()(IRtdNotify& notify)
      {
        return getWorkerThreadPool()->push(
          [=, &notify](int /*threadId*/)
          {
            py::gil_scoped_acquire gilAcquired;

            PyErr_Clear();

            auto kwargs = py::reinterpret_borrow<py::object>(_kwargs);
            kwargs[THREAD_CONTEXT_TAG] = RtdNotifier(notify);

            ExcelObj result;
            invokePyFunction(result, _info, _args, kwargs.ptr());
            notify.publish(std::move(result));
          }
        );
      }

      bool operator==(const AsyncTask& that) const
      {
        py::gil_scoped_acquire gilAcquired;

        auto args = py::reinterpret_borrow<py::tuple>(_args);
        auto kwargs = py::reinterpret_borrow<py::dict>(_kwargs);
        auto that_args = py::reinterpret_borrow<py::tuple>(that._args);
        auto that_kwargs = py::reinterpret_borrow<py::dict>(that._kwargs);

        if (args.size() != that_args.size() 
          || kwargs.size() != that_kwargs.size())
          return false;

        for (auto i = args.begin(), j = that_args.begin();
          i != args.end();
          ++i, ++j)
        {
          if (!i->equal(*j))
            return false;
        }
        for (auto i = kwargs.begin(); i != kwargs.end(); ++i)
        {
          if (!i->first.equal(py::str(THREAD_CONTEXT_TAG))
            && !i->second.equal(that_kwargs[i->first]))
            return false;
        }
        return true;
      }
    };

    ExcelObj* pythonRtdCallback(
      PyFuncInfo* info, 
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        // TODO: consider argument capture and equality check under c++
        PyObject *argsP, *kwargsP;
        {
          py::gil_scoped_acquire gilAcquired;

          auto[args, kwargs] = convertArgsToPython(info, xlArgs);
          if (kwargs.is_none())
            kwargs = py::dict();

          // Add this here so that dict sizes for running and newly 
          // created tasks match
          kwargs[THREAD_CONTEXT_TAG] = py::none();

          // Need to drop pybind links before capturing in lambda otherwise the destructor
          // is called at some random time after losing the GIL and it crashes.
          argsP = args.release().ptr();
          kwargsP = kwargs.release().ptr();
        }

        auto value = rtdAsync(
          std::make_shared<RtdAsyncTask<AsyncTask>>(info, argsP, kwargsP));
        return returnValue(value ? *value : CellError::NA);
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

    shared_ptr<const FuncSpec> createSpec(const shared_ptr<PyFuncInfo>& funcInfo)
    {
      shared_ptr<const FuncSpec> spec;
      if (funcInfo->info->options & FuncInfo::ASYNC)
      {
        getWorkerThreadPool(); // Ensure initialised
        if (funcInfo->isRtdAsync)
          XLO_THROW("Cannot specify async registration with Rtd async");

        spec.reset(new AsyncCallbackSpec(funcInfo->info, &pythonAsyncCallback, funcInfo));
      }
      else if (funcInfo->isRtdAsync)
      {
        getWorkerThreadPool();// Ensure initialised
        spec.reset(new CallbackSpec(funcInfo->info, &pythonRtdCallback, funcInfo));
      }
      else
        spec.reset(new CallbackSpec(funcInfo->info, &pythonCallback, funcInfo));

      return spec;
    }

    void handleFileChange(
      const wchar_t* dirName,
      const wchar_t* fileName,
      const Event::FileAction action);

    class RegisteredModule : public FileSource
    {
    public:
      RegisteredModule(
        const wstring& modulePath,
        const wchar_t* workbookName)
        : FileSource(modulePath.c_str(), workbookName)
      {
        auto path = fs::path(modulePath);
        _fileWatcher = std::static_pointer_cast<const void>
          (Event::DirectoryChange(path.remove_filename()).bind(handleFileChange));
        if (workbookName)
          _workbookName = workbookName;
      }

      void registerPyFuncs(
        PyObject* pyModule,
        const vector<shared_ptr<PyFuncInfo>>& functions)
      {
        _module = pyModule;
        vector<shared_ptr<const FuncSpec>> nonLocal;
        vector<shared_ptr<const FuncInfo>> funcInfo;
        vector<ExcelFuncObject> funcs;

        for (auto& f : functions)
        {
          if (_workbookName.empty())
            f->isLocalFunc = false;
          if (!f->isLocalFunc)
            nonLocal.push_back(createSpec(f));
          else
          {
            funcInfo.push_back(f->info);
            funcs.push_back([f](const FuncInfo&, const ExcelObj** args)
            {
              return pythonCallback(f.get(), args);
            });
          }
        }

        registerFuncs(nonLocal);
        for (auto& f : nonLocal)
          XLO_ERROR(L"Registration failed for: {0}", f->name());

        if (!funcInfo.empty())
        {
          if (_workbookName.empty())
            XLO_THROW("Local functions found without workbook specification");
          registerLocal(funcInfo, funcs);
        }
      }

      // TODO: Is it possible the module will be unloaded?
      // We don't really want to have to deal with the GIL
      PyObject* pyModule() const { return _module; }

      const wchar_t* workbookName() const
      {
        return _workbookName.empty() ? nullptr : _workbookName.c_str();
      }

    private:
      shared_ptr<const void> _fileWatcher;
      wstring _workbookName;
      PyObject* _module = Py_None;
    };

    std::shared_ptr<RegisteredModule>
      FunctionRegistry::addModule(
        AddinContext* context,
        const std::wstring& modulePath,
        const wchar_t* workbookName)
    {
      auto[fileCtx, inserted] = context->tryAdd<RegisteredModule>(
        modulePath.c_str(), modulePath, workbookName);
      return fileCtx;
    }

    void handleFileChange(
      const wchar_t* dirName,
      const wchar_t* fileName,
      const Event::FileAction action)
    {
      const auto filePath = (fs::path(dirName) / fileName).wstring();
      
      auto[foundSource, foundAddin] = FileSource::findFileContext(filePath.c_str());
      
      // If no active filecontext is found, then exit. Note that findFileContext
      // will check if a linked workbook is still open 
      if (!foundSource)
        return;

      switch (action)
      {
      case Event::FileAction::Modified:
      {
        XLO_INFO(L"Module '{0}' modified, reloading.", filePath);
        // TODO: can we be sure about this context setting?
        theCurrentContext = foundAddin.get();
        // Our file source must be of type RegisteredModule so the cast is safe
        auto& pySource = static_cast<RegisteredModule&>(*foundSource);
        // Rescan the module, passing in the module handle if it exists
        py::gil_scoped_acquire get_gil;
        scanModule(
          pySource.pyModule() != Py_None 
            ? PyBorrow<py::module>(pySource.pyModule())
            : py::wstr(filePath),
          pySource.workbookName());
        // Set the addin context back. Not exeception safe clearly.
        theCurrentContext = theCoreContext;
        break;
      }
      case Event::FileAction::Delete:
      {
        XLO_INFO(L"Module '{0}' deleted/renamed, removing functions.", filePath);
        FileSource::deleteFileContext(foundSource);
        break;
      }
      }
    }

    void registerFunctions(
      const py::object& moduleHandle,
      const vector<shared_ptr<PyFuncInfo>>& functions)
    {
      wstring modulePath;
      {
        py::gil_scoped_acquire get_gil;
        modulePath = moduleHandle.attr("__file__").cast<wstring>();
      }
      auto mod = FunctionRegistry::addModule(
        theCurrentContext, modulePath);
      mod->registerPyFuncs(moduleHandle.ptr(), functions);
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
          .def_readwrite("local", &PyFuncInfo::isLocalFunc)
          .def_readwrite("rtd_async", &PyFuncInfo::isRtdAsync);

        py::class_<AsyncNotifier>(mod, "AsyncNotifier")
          .def("cancelled", &AsyncNotifier::cancelled);

        py::class_<RtdNotifier>(mod, "RtdNotifier")
          .def("cancelled", &RtdNotifier::cancelled);

        mod.add_object("ASYNC_CONTEXT_TAG", py::str(THREAD_CONTEXT_TAG));
        mod.def("register_functions", &registerFunctions);
      });
    }
  }
}