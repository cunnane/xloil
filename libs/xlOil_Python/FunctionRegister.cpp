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
        return returnValue(e);
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
        return returnValue(e.what());
      }
      catch (...)
      {
        return returnValue(CellError::Null);
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
      decltype(GetTickCount64()) _startTime;
    };

    static ctpl::thread_pool* thePythonWorkerThread = nullptr;

    static auto WorkerThreadDeleter = Event_PyBye().bind([]
    {
      if (thePythonWorkerThread)
        delete thePythonWorkerThread;
    });

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
        XLO_WARN(e.what());
        asyncReturn(*asyncHandle, ExcelObj(e.what()));
      }
      catch (...)
      {
        XLO_WARN("Async unknown error");
        asyncReturn(*asyncHandle, ExcelObj(CellError::Value));
      }
    }

    shared_ptr<const FuncSpec> createSpec(const shared_ptr<PyFuncInfo>& funcInfo)
    {
      shared_ptr<const FuncSpec> spec;
      if (funcInfo->info->options & FuncInfo::ASYNC)
      {
        if (!thePythonWorkerThread)
          thePythonWorkerThread = new ctpl::thread_pool(1);

        spec.reset(new AsyncCallbackSpec(funcInfo->info, &pythonAsyncCallback, funcInfo));
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
        : FileSource(modulePath.c_str())
      {
        const auto path = fs::path(modulePath);
        _fileWatcher = std::static_pointer_cast<const void>
          (Event::DirectoryChange(fs::path(path).remove_filename()).bind(handleFileChange));
        if (workbookName)
          _workbookName = workbookName;
      }

      void registerPyFuncs(
        const py::module& pyModule,
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
          registerLocal(_workbookName.c_str(), funcInfo, funcs);
        }
      }

      const py::module& pyModule() const { return _module; }

      const wchar_t* workbookName() const
      {
        return _workbookName.empty() ? nullptr : _workbookName.c_str();
      }

    private:
      shared_ptr<const void> _fileWatcher;
      wstring _workbookName;
      py::module _module;
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
      auto filePath = (fs::path(dirName) / fileName).wstring();
      auto[foundSource, foundAddin] = FileSource::findFileContext(filePath.c_str());
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
        scanModule(
          !pySource.pyModule().is_none() ? pySource.pyModule() : py::wstr(filePath),
          pySource.workbookName());
        // Set the addin context back. Not exeception safe clearly.
        theCurrentContext = theCoreContext;
        break;
      }
      case Event::FileAction::Delete:
      {
        XLO_INFO(L"Module '{0}' deleted, removing functions.", filePath);
        FileSource::deleteFileContext(foundSource);
        break;
      }
      }
    }

    void registerFunctions(
      const py::object& moduleHandle,
      const vector<shared_ptr<PyFuncInfo>>& functions)
    {
      auto mod = FunctionRegistry::addModule(
        theCurrentContext, moduleHandle.attr("__file__").cast<wstring>());
      mod->registerPyFuncs(moduleHandle, functions);
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

        mod.def("register_functions", &registerFunctions);
      });
    }
  }
}