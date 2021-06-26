#include "FunctionRegister.h"
#include "PyCoreModule.h"
#include "Main.h"
#include "BasicTypes.h"
#include "Dictionary.h"
#include "ReadSource.h"
#include "FunctionRegister.h"
#include "AsyncFunctions.h"
#include "PyEvents.h"
#include <xloil/StaticRegister.h>
#include <xloil/DynamicRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/Caller.h>
#include <xloil/RtdServer.h>
#include <xlOil/ExcelApp.h>
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
    constexpr wchar_t* XLOPY_ANON_SOURCE = L"PythonFuncs";
    constexpr char* XLOPY_CLEANUP_FUNCTION = "_xloil_unload";

    void setFuncType(PyFuncInfo& info, const string& type, bool isVolatile)
    {
      unsigned base = isVolatile ? FuncInfo::VOLATILE : 0;
      if (type == "macro")
      {
        info.setFuncOptions(FuncInfo::MACRO_TYPE & base);
      }
      else if (type == "threaded")
      {
        info.setFuncOptions(FuncInfo::THREAD_SAFE & base);
      }
      else if (type == "rtd")
      {
        info.isRtdAsync = true;
      }
      else if (type == "async")
      {
        info.isAsync = true;
      }
    }

    PyFuncInfo::PyFuncInfo(
      const std::wstring& name,
      const pybind11::function& func,
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
    {
      _info->name = name;
      _info->help = help;
      _info->category = category;
      
      if (!func.ptr() || func.is_none())
        XLO_THROW(L"No python function specified for {0}", name);

      setFuncType(*this, features, isVolatile);

      _info->args.resize(numArgs + (isAsync ? 1 : 0));
      if (isAsync)
        _info->args[0] = FuncArg(nullptr, nullptr, FuncArg::AsyncHandle);

      for (auto i = 0u; i < numArgs; ++i)
        _args.push_back(PyFuncArg(_info, i + (isAsync ? 1 : 0)));
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

    pair<py::tuple, py::object> PyFuncInfo::convertArgs(const ExcelObj** xlArgs) const
    {
      auto nArgs = _args.size() - (_hasKeywordArgs ? 1u : 0);
      auto pyArgs = PySteal<py::tuple>(PyTuple_New(nArgs));

      // TODO: is it worth having a enum switch to convert primitive types rather than a v-call
      for (auto i = 0u; i < nArgs; ++i)
      {
        try
        {
          auto* defaultValue = _args[i].getDefault().ptr();
          auto* pyObj = (*_args[i].converter)(*xlArgs[i], defaultValue);
          PyTuple_SET_ITEM(pyArgs.ptr(), i, pyObj);
        }
        catch (const std::exception& e)
        {
          // We give the arg number 1-based as it's more natural
          XLO_THROW(L"Error in arg {1} '{0}': {2}",
            _args[i].arg.name, std::to_wstring(i + 1), utf8ToUtf16(e.what()));
        }
      }
      if (_hasKeywordArgs)
      {
        auto kwargs = PySteal<py::dict>(readKeywordArgs(*xlArgs[nArgs]));
        return make_pair(pyArgs, kwargs);
      }
      else
        return make_pair(pyArgs, py::none());
    }

    void PyFuncInfo::invoke(PyObject* args, PyObject* kwargs) const
    {
      PyObject* ret;
      if (kwargs != Py_None)
        ret = PyObject_Call(_func.ptr(), args, kwargs);
      else
        ret = PyObject_CallObject(_func.ptr(), args);
      if (!ret)
        throw py::error_already_set();
    }

    void PyFuncInfo::invoke(
      ExcelObj& result, 
      PyObject* args, 
      PyObject* kwargs) const noexcept
    {
      try
      {
        py::object ret;
        if (kwargs != Py_None)
          ret = PySteal<py::object>(PyObject_Call(_func.ptr(), args, kwargs));
        else
          ret = PySteal<py::object>(PyObject_CallObject(_func.ptr(), args));

        result = returnConverter
          ? (*returnConverter)(*ret.ptr())
          : FromPyObj()(ret.ptr());
      }
      catch (const py::error_already_set& e)
      {
        Event_PyUserException().fire(e.type(), e.value(), e.trace());
        result = e.what();
      }
      catch (const std::exception& e)
      {
        result = e.what();
      }
    }

    ExcelObj* pythonCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        auto[args, kwargs] = info->convertArgs(xlArgs);

        static ExcelObj result; // Ok since we have the GIL
        info->invoke(result, args.ptr(), kwargs.ptr());

        // It's not safe to return the static object if the function
        // is being multi-threaded by Excel as we can't control when
        // Excel will read the result.
        if (info->isThreadSafe())
          return returnValue(result);
        else
          return &result;
      }
      catch (const py::error_already_set& e)
      {
        Event_PyUserException().fire(e.type(), e.value(), e.trace());
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


    shared_ptr<const WorksheetFuncSpec> createSpec(const shared_ptr<const PyFuncInfo>& funcInfo)
    {
      auto info = funcInfo->info();
      if (funcInfo->isAsync)
        return make_shared<DynamicSpec>(info, &pythonAsyncCallback, funcInfo);
      else if (funcInfo->isRtdAsync)
        return make_shared<DynamicSpec>(info, &pythonRtdCallback, funcInfo);
      else
        return make_shared<DynamicSpec>(info, &pythonCallback, funcInfo);
    }

    class WatchedSource : public FileSource
    {
    public:
      WatchedSource(
        const wchar_t* sourceName,
        const wchar_t* linkedWorkbook = nullptr)
        : FileSource(sourceName, linkedWorkbook)
      {
        auto path = fs::path(sourceName);
        auto dir = path.remove_filename();
        if (!dir.empty())
          _fileWatcher = Event::DirectoryChange(dir)->bind(
            [this](auto dir, auto file, auto act)
        {
          handleDirChange(dir, file, act);
        });

        if (linkedWorkbook)
          _workbookWatcher = Event::WorkbookAfterClose().bind([this](auto wb) { handleClose(wb); });
      }

      virtual void reload() = 0;

    private:
      shared_ptr<const void> _fileWatcher;
      shared_ptr<const void> _workbookWatcher;

      void handleClose(
        const wchar_t* wbName)
      {
        if (_wcsicmp(wbName, linkedWorkbook().c_str()) == 0)
          FileSource::deleteFileContext(shared_from_this());
      }

      void handleDirChange(
        const wchar_t* dirName,
        const wchar_t* fileName,
        const Event::FileAction action)
      {
        if (_wcsicmp(fileName, sourceName()) != 0)
          return;
        
        excelRunOnMainThread([
            this,
            dirStr = wstring(dirName),
            fileStr = wstring(fileName),
            action]()
          {
            const auto filePath = (fs::path(dirStr) / fileStr).wstring();

            // Directories should match as our directory watch listener only checks
            // the specified directory
            assert(_wcsicmp(filePath.c_str(), sourcePath().c_str()) == 0);

            switch (action)
            {
            case Event::FileAction::Modified:
            {
              XLO_INFO(L"Module '{0}' modified, reloading.", filePath);
              reload();
              break;
            }
            case Event::FileAction::Delete:
            {
              XLO_INFO(L"Module '{0}' deleted/renamed, removing functions.", filePath);
              FileSource::deleteFileContext(shared_from_this());
              break;
            }
            }
          }, ExcelRunQueue::ENQUEUE);
      }
    };

    //TODO: Refactor Python FileSource
    // It might be better for lifetime management if the whole FileSource interface was exposed
    // via the core, then a reference to the FileSource can be held and closed by the module 
    // itself
    class RegisteredModule : public WatchedSource
    {
    public:
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
          // TODO: cancel running async tasks?
          py::gil_scoped_acquire getGil;

          // Call module cleanup function
          auto thisMod = PyBorrow<py::module>(_module);
          if (py::hasattr(thisMod, XLOPY_CLEANUP_FUNCTION))
            thisMod.attr(XLOPY_CLEANUP_FUNCTION)();
         
          auto success = unloadModule(thisMod);

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
        PyObject* pyModule,
        const vector<shared_ptr<PyFuncInfo>>& functions)
      {
        // Note we don't increment the ref-counter for the module to 
        // simplify our destructor
        // TODO: this is not safe!
        _module = pyModule;
        vector<shared_ptr<const WorksheetFuncSpec>> nonLocal;
        vector<shared_ptr<const FuncInfo>> funcInfo;
        vector<DynamicExcelFunc<>> funcs;

        for (auto& f : functions)
        {
          if (!_linkedWorkbook)
            f->isLocalFunc = false;
          if (!f->isLocalFunc)
            nonLocal.push_back(createSpec(f));
          else
          {
            funcInfo.push_back(f->info());
            if (f->isRtdAsync)
            {
              funcs.emplace_back([f](const FuncInfo&, const ExcelObj** args)
              {
                return pythonRtdCallback(f.get(), args);
              });
            }
            else
            {
              funcs.emplace_back([f](const FuncInfo&, const ExcelObj** args)
              {
                return pythonCallback(f.get(), args);
              });
            }
          }
        }

        registerFuncs(nonLocal);

        if (!funcInfo.empty())
        {
          if (!_linkedWorkbook)
            XLO_THROW("Local functions found without workbook specification");
          registerLocal(funcInfo, funcs);
        }
      }

      void reload()
      {
        // TODO: can we be sure about this context setting?
        // 
        auto[source, addin] = FileSource::findFileContext(sourcePath().c_str());
        if (source.get() != this)
          XLO_THROW(L"Error reloading '{0}': source ptr mismatch", sourcePath());

        auto currentContext = theCurrentContext;
        theCurrentContext = addin.get();

        // Rescan the module, passing in the module handle if it exists
        py::gil_scoped_acquire get_gil;
        py::object moduleHandle;
        if (_module != Py_None)
        {
          moduleHandle = PyBorrow<py::object>(_module);
          moduleHandle.cast<py::module>().reload();
        }
        else
        {
          moduleHandle = loadModuleFromFile(sourcePath().c_str(), linkedWorkbook().c_str());
        }
        scanModule(moduleHandle);

        // Set the addin context back. TODO: Not exception safe clearly.
        theCurrentContext = currentContext;
      }

    private:
      bool _linkedWorkbook;
      PyObject* _module = Py_None;
    };

    std::shared_ptr<RegisteredModule>
      FunctionRegistry::addModule(
        AddinContext* context,
        const std::wstring& modulePath,
        const wchar_t* workbookName)
    {
      auto[source, addin] = FileSource::findFileContext(modulePath.c_str());
      if (source)
        return std::static_pointer_cast<RegisteredModule>(source);

      auto fileSrc = make_shared<RegisteredModule>(modulePath, workbookName);
      context->addSource(fileSrc);
      return fileSrc;
    }

    void registerFunctions(
      const py::object& moduleHandle,
      const vector<shared_ptr<PyFuncInfo>>& functions)
    {
      // Called from python so we have the GIL
      // The "null" module handle is used by jupyter
      const auto modulePath = !moduleHandle.is_none()
        ? moduleHandle.attr("__file__").cast<wstring>()
        : L"";

      py::gil_scoped_release releaseGil;

      auto mod = FunctionRegistry::addModule(
        theCurrentContext, modulePath, nullptr);
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
          .def(py::init<wstring, py::function, unsigned, string, wstring, wstring, bool, bool, bool>(),
            py::arg("name"),
            py::arg("func"),
            py::arg("nargs"),
            py::arg("features") = py::none(),
            py::arg("help") = "",
            py::arg("category") = "",
            py::arg("local") = true,
            py::arg("volatile") = false,
            py::arg("has_kwargs") = false)
          .def_property("return_converter", &PyFuncInfo::getReturnConverter, &PyFuncInfo::setReturnConverter)
          .def_property_readonly("args", &PyFuncInfo::args);

        mod.def("register_functions", &registerFunctions);
        mod.def("deregister_functions", &deregisterFunctions);
      });
    }
  }
}