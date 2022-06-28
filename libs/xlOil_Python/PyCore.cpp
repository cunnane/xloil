#include "PyCore.h"
#include "PyHelpers.h"
#include "PyEvents.h"
#include "EventLoop.h"
#include "PyFuture.h"
#include "TypeConversion/Numpy.h"
#include <TypeConversion/BasicTypes.h>
#include <xlOil/ExcelThread.h>
#include <xloil/Caller.h>
#include <xloil/State.h>
#include <xlOil/StringUtils.h>
#include <map>

using std::shared_ptr;
using std::vector;
using std::wstring;
using std::string;
namespace py = pybind11;
using std::make_pair;

namespace xloil 
{
  namespace Python 
  {
    using BinderFunc = std::function<void(pybind11::module&)>;

    PyTypeObject* cellErrorType;
    PyObject*     comBusyException;
    PyObject*     cannotConvertException;
    shared_ptr<const IPyToExcel> theCustomReturnConverter = nullptr;

    namespace
    {
      auto cleanupGlobals = Event_PyBye().bind([] {
        theCustomReturnConverter.reset();
      });
      void initialiseCore(py::module& mod);
    }

    class BinderRegistry
    {
    public:
      static BinderRegistry& get() {
        static BinderRegistry instance;
        return instance;
      }

      auto add(BinderFunc f, size_t priority)
      {
        return theFunctions.insert(make_pair(priority, f));
      }

      void bindAll(py::module& mod)
      {
        std::for_each(theFunctions.rbegin(), theFunctions.rend(),
          [&mod](auto f) { f.second(mod); });
      }
    private:
      BinderRegistry() {}
      std::multimap<size_t, BinderFunc> theFunctions;
    };

    PyObject* buildInjectedModule()
    {
      auto mod = py::module::create_extension_module(
        theInjectedModuleName, nullptr, new PyModuleDef());
      initialiseCore(mod);
      BinderRegistry::get().bindAll(mod);
      return mod.release().ptr();
    }

    PYBIND11_MODULE(XLO_PROJECT_NAME, mod)
    {
      mod.doc() = R"(
        The Python plugin for xlOil primarily allows creation of Excel functions and macros 
        backed by Python code. In addition it offers full control over GUI objects and an 
        interface for Excel automation: driving the application in code.

        See the documentation at https://xloil.readthedocs.io
      )";

      initialiseCore(mod);
      BinderRegistry::get().bindAll(mod);
    }

    int addBinder(std::function<void(pybind11::module&)> binder)
    {
      BinderRegistry::get().add(binder, 1);
      return 0;
    }

    namespace
    {
      auto runLater(
        const py::object& callable, 
        const unsigned delay, 
        const unsigned retryPause, 
        wstring& api)
      {
        int flags = 0;
        toLower(api);
        if (api.empty() || api.find(L"com") != wstring::npos)
          flags |= ExcelRunQueue::COM_API;
        if (api.find(L"xll") != wstring::npos)
          flags |= ExcelRunQueue::XLL_API;
        if (retryPause == 0)
          flags |= ExcelRunQueue::NO_RETRY;
        return PyFuture<PyObject*>(runExcelThread([callable=PyObjectHolder(callable)]()
          {
            py::gil_scoped_acquire getGil;
            try
            {
              return callable().release().ptr();
            }
            catch (py::error_already_set& err)
            {
              if (err.matches(comBusyException))
                throw ComBusyException();
              throw;
            }
          },
          flags,
          delay,
          retryPause));
      }

      void setReturnConverter(const shared_ptr<const IPyToExcel>& conv)
      {
        theCustomReturnConverter = conv;
      }

      struct CannotConvert {};

      auto cellErrorSymbol(CellError e)
      {
        auto wstr = enumAsWCString(e);
        string str;
        for (auto c = wstr; *c != L'0'; ++c)
        {
          if (*c != L'#' && *c != L'/' && *c != L'?' && *c != L'!')
            str.push_back((char)*c);
        }
        return str;
      }

      void initialiseCore(pybind11::module& mod)
      {
        XLO_DEBUG("Python importing numpy");
        if (!importNumpy())
          throw py::error_already_set();

        // Bind the two base classes for python converters
        py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel")
          .def("__call__",
            [](const IPyFromExcel& /*self*/, const py::object& /*arg*/)
            {
              XLO_THROW("Internal IPyFromExcel converters cannot be called from python");
            });

        py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

        mod.def("set_return_converter", setReturnConverter);

        mod.def("in_wizard", &inFunctionWizard,
          R"(
          Returns true if the function is being invoked from the function wizard : costly functions should"
          exit in this case to maintain UI responsiveness.Checking for the wizard is itself not cheap, so"
          use this sparingly.
          )");

        mod.def("excel_callback",
          &runLater,
          R"(
          Schedules a callback to be run in the main thread. Much of the COM API in unavailable
          during the calc cycle, in particular anything which involves writing to the sheet.
          Returns a future which can be awaited.

          Parameters
          ----------

          func: callable
          A callable which takes no arguments and returns nothing

          retry : int
          Millisecond delay between retries if Excel's COM API is busy, e.g. a dialog box
          is open or it is running a calc cycle.If zero, does no retry

          wait : int
          Number of milliseconds to wait before first attempting to run this function

          api : str
          Specify 'xll' or 'com' or both to indicate which APIs the call requires.
          The default is 'com': 'xll' would only be required in rare cases.
          )",
          py::arg("func"),
          py::arg("wait") = 0,
          py::arg("retry") = 500,
          py::arg("api") = "");

        py::class_<Environment::ExcelProcessInfo>(mod, "ExcelState", 
          R"(
          Gives information about the Excel application, in particular the handles required
          to interact with Excel via the Win32 API.
          )")
          .def_readonly("version", &Environment::ExcelProcessInfo::version, 
            "Excel major version")
          .def_readonly("hinstance", &Environment::ExcelProcessInfo::hInstance,
            "Excel HINSTANCE")
          .def_readonly("hwnd", &Environment::ExcelProcessInfo::hWnd,
            "Excel main window handle(as an int)")
          .def_readonly("main_thread_id", &Environment::ExcelProcessInfo::mainThreadId,
            "Excel's main thread ID");

        mod.def("excel_state", Environment::excelProcess);

        comBusyException = py::register_exception<ComBusyException>(mod, "ComBusyError").ptr();

        {
          auto e = py::exception<CannotConvert>(mod, "CannotConvert");
          e.doc() = R"(
            Should be thrown by a converter when it is unable to handle the 
            provided type.  In a return converter it may not indicate a fatal 
            condition, as xlOil will fallback to another converter.
          )";
          cannotConvertException = e.ptr();
        }

        {
          // Bind CellError type to xloil::CellError enum
          auto eType = py::enum_<CellError>(mod, "CellError", 
            R"(
              Enum-type class which represents an Excel error condition of the 
              form `#N/A!`, `#NAME!`, etc passed as a function argument. If a 
              function argument does not specify a type (e.g. int, str) it may be passed 
              a CellError, which it can handle based on the error condition.
            )");

          for (auto e : theCellErrors)
            eType.value(cellErrorSymbol(e).c_str(), e);

          cellErrorType = (PyTypeObject*)eType.ptr();
        }

        mod.def("get_event_loop", [](const wchar_t* addin) { findAddin(addin).thread->loop(); });
      }
    }
} }