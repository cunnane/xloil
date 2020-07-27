#include "FunctionRegister.h"
#include "InjectedModule.h"
#include "Main.h"
#include "BasicTypes.h"
#include "Dictionary.h"
#include "File.h"
#include "FunctionRegister.h"
#include "AsyncFunctions.h"
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Caller.h>
#include <xloil/RtdServer.h>
#include <xloil/ThreadControl.h>
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
namespace py = pybind11;
using namespace pybind11::literals;

namespace xloil
{
  namespace Python
  {
    constexpr wchar_t* PYTHON_ANON_SOURCE = L"PythonFuncs";

    PyFuncInfo::PyFuncInfo(
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
      if (!info)
        XLO_THROW("No function info specified in func registration");
      if (!func.ptr() || func.is_none())
        XLO_THROW(L"No python function specified for {0}", info->name);
    }

    void PyFuncInfo::setArgTypeDefault(size_t i, shared_ptr<IPyFromExcel> converter, py::object defaultVal)
    {
      argConverters[i] = std::make_pair(converter, defaultVal);
    }

    void PyFuncInfo::setArgType(size_t i, shared_ptr<IPyFromExcel> converter)
    {
      argConverters[i] = std::make_pair(converter, py::object());
    }

    void PyFuncInfo::setFuncOptions(int val)
    {
      info->options = val;
    }

    pair<py::tuple, py::object> PyFuncInfo::convertArgs(const ExcelObj** xlArgs)
    {
      auto nArgs = argConverters.size();
      auto pyArgs = PySteal<py::tuple>(PyTuple_New(nArgs));

      // TODO: is it worth having a enum switch to convert primitive types rather than a v-call
      for (auto i = 0u; i < nArgs; ++i)
      {
        try
        {
          auto* def = argConverters[i].second.ptr();
          auto* pyObj = (*argConverters[i].first)(*xlArgs[i], def);
          PyTuple_SET_ITEM(pyArgs.ptr(), i, pyObj);
        }
        catch (const std::exception& e)
        {
          // TODO: could we explain what type is required?
          // We give the arg number 1-based as it's more natural
          XLO_THROW(L"Error reading '{0}' arg #{1}: {2}",
            info->args[i].name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
        }
      }
      if (hasKeywordArgs)
      {
        auto kwargs = PySteal<py::dict>(readKeywordArgs(*xlArgs[nArgs]));
        return make_pair(pyArgs, kwargs);
      }
      else
        return make_pair(pyArgs, py::none());
    }

    void PyFuncInfo::invoke(PyObject* args, PyObject* kwargs)
    {
      PyObject* ret;
      if (kwargs != Py_None)
        ret = PyObject_Call(func.ptr(), args, kwargs);
      else
        ret = PyObject_CallObject(func.ptr(), args);
      if (!ret)
        throw py::error_already_set();
    }

    void PyFuncInfo::invoke(ExcelObj& result, PyObject* args, PyObject* kwargs) noexcept
    {
      try
      {
        py::object ret;
        if (kwargs != Py_None)
          ret = PySteal<py::object>(PyObject_Call(func.ptr(), args, kwargs));
        else
          ret = PySteal<py::object>(PyObject_CallObject(func.ptr(), args));

        result = returnConverter
          ? (*returnConverter)(*ret.ptr())
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

        auto[args, kwargs] = info->convertArgs(xlArgs);

        static ExcelObj result; // Ok since python is single threaded
        info->invoke(result, args.ptr(), kwargs.ptr());
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


    shared_ptr<const FuncSpec> createSpec(const shared_ptr<PyFuncInfo>& funcInfo)
    {
      shared_ptr<const FuncSpec> spec;
      if (funcInfo->info->options & FuncInfo::ASYNC)
      {
        if (funcInfo->isRtdAsync)
          XLO_THROW("Cannot specify async registration with Rtd async");

        spec.reset(new AsyncCallbackSpec(funcInfo->info, &pythonAsyncCallback, funcInfo));
      }
      else if (funcInfo->isRtdAsync)
      {
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
        : FileSource(
            modulePath.empty() ? PYTHON_ANON_SOURCE : modulePath.c_str(), 
            workbookName)
      {
        if (!modulePath.empty())
        {
          auto path = fs::path(modulePath);
          _fileWatcher = std::static_pointer_cast<const void>
            (Event::DirectoryChange(path.remove_filename()).bind(handleFileChange));
        }
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
      auto[fileCtx, inserted] 
        = context->tryAdd<RegisteredModule>(
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

        // Set the addin context back. TODO: Not exeception safe clearly.
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
      // Called from python so we have the GIL
      const auto modulePath = !moduleHandle.is_none()
        ? moduleHandle.attr("__file__").cast<wstring>()
        : L"";

      py::gil_scoped_release releaseGil;

      auto mod = FunctionRegistry::addModule(
        theCurrentContext, modulePath);
      mod->registerPyFuncs(moduleHandle.ptr(), functions);
    }
    void deregisterFunctions(
      const py::object& moduleHandle,
      const py::object& functionNames)
    {
      // Called from python so we have the GIL

      const auto modulePath = !moduleHandle.is_none()
        ? moduleHandle.attr("__file__").cast<wstring>()
        : L"";
      
      auto[foundSource, foundAddin] = FileSource::findFileContext(
        modulePath.empty() ? PYTHON_ANON_SOURCE : modulePath.c_str());

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

      for (auto& func : funcNames)
        foundSource->deregister(func);
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

        // TODO: Both these classes have terrible names...can we improve them?
        py::class_<PyFuncInfo, shared_ptr<PyFuncInfo>>(mod, "FuncHolder") 
          .def(py::init<const shared_ptr<FuncInfo>&, const py::function&, bool>())
          .def("set_arg_type", &PyFuncInfo::setArgType, py::arg("i"), py::arg("arg_type"))
          .def("set_arg_type_defaulted", &PyFuncInfo::setArgTypeDefault, py::arg("i"), py::arg("arg_type"), py::arg("default"))
          .def("set_opts", &PyFuncInfo::setFuncOptions, py::arg("flags"))
          .def_readwrite("local", &PyFuncInfo::isLocalFunc)
          .def_readwrite("rtd_async", &PyFuncInfo::isRtdAsync);

        mod.def("register_functions", &registerFunctions);
        mod.def("deregister_functions", &deregisterFunctions);
      });
    }
  }
}