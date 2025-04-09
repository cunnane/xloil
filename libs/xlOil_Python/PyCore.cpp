#include "PyCore.h"
#include "PyHelpers.h"
#include "PyEvents.h"
#include "EventLoop.h"
#include "PyFuture.h"
#include "TypeConversion/Numpy.h"
#include <TypeConversion/BasicTypes.h>
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
using ExcelProcessInfo = xloil::Environment::ExcelProcessInfo;

namespace xloil 
{
  namespace Python 
  {
    PyTypeObject* theCellErrorType;
    PyObject*     cannotConvertException;
    PyTypeObject* theExcelObjType;

    bool isErrorType(const PyObject* obj)
    {
      return Py_TYPE(obj) == theCellErrorType;
    }

    bool isExcelObjType(const PyObject* obj)
    {
      return Py_TYPE(obj) == theExcelObjType;
    }

    namespace
    {
      using BinderFunc = std::function<void(pybind11::module&)>;

      void initialiseCore(py::module& mod);
   
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
    }

    PyObject* buildInjectedModule()
    {
      auto mod = py::module::create_extension_module(
        theInjectedModuleName, nullptr, new PyModuleDef());
      initialiseCore(mod);
      BinderRegistry::get().bindAll(mod);
      return mod.release().ptr();
    }

    // This unfortunate block of code is a copy of PYBIND11_MODULE with the name of 
    // the module tweaked. This allows the module name to be consistent across the 
    // various xloil_PythonXX.pyd implemenations which reduces surprise and makes 
    // the documentation nicer

#define XLO_NAMED_MODULE(name, variable, ModuleName)                                                           \
    static ::pybind11::module_::module_def PYBIND11_CONCAT(pybind11_module_def_, name); \
    static void PYBIND11_CONCAT(pybind11_init_, name)(::pybind11::module_ &);           \
    PYBIND11_PLUGIN_IMPL(name) {                                                        \
        PYBIND11_CHECK_PYTHON_VERSION                                                   \
        PYBIND11_ENSURE_INTERNALS_READY                                                 \
        auto m = ::pybind11::module_::create_extension_module(                          \
            ModuleName, nullptr, &PYBIND11_CONCAT(pybind11_module_def_, name));         \
        try {                                                                           \
            PYBIND11_CONCAT(pybind11_init_, name)(m);                                   \
            return m.ptr();                                                             \
        }                                                                               \
        PYBIND11_CATCH_INIT_EXCEPTIONS                                                  \
    }                                                                                   \
    void PYBIND11_CONCAT(pybind11_init_, name)(::pybind11::module_ & (variable))


    XLO_NAMED_MODULE(XLO_PROJECT_NAME, mod, theInjectedModuleName)
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

    int addBinder(std::function<void(pybind11::module&)> binder, size_t priority)
    {
      BinderRegistry::get().add(binder, priority);
      return 0;
    }

    namespace 
    {
      struct CannotConvert {};

      /// <summary>
      /// Gets rid of any #, /, ? or ! chars from the cell errors to 
      /// produce a valid python symbol.
      /// </summary>
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

        importDatetime();

        theExcelObjType = (PyTypeObject*)py::class_<ExcelObj>(
          mod, 
          "_RawExcelValue"
          R"(
            Wrapper for an already-converted Excel value ready to be returned directly
            to Excel. For internal use in type converters.
          )").ptr();

        // Bind the two base classes for python converters
        py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel")
          .def("__call__",
            [](const IPyFromExcel& /*self*/, const py::object& /*arg*/)
            {
              XLO_THROW("Internal IPyFromExcel converters cannot be called from python");
            })
          .def("__str__", [](const IPyFromExcel& self) { return self.name(); });

        py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel")
          .def("__str__", [](const IPyToExcel& self) { return self.name(); })
          .def("__call__", &IPyToExcel::operator());

        mod.def("in_wizard", &inFunctionWizard,
          R"(
            Returns true if the function is being invoked from the function wizard : costly functions 
            should exit in this case to maintain UI responsiveness.  Checking for the wizard is itself 
            not cheap, so use this sparingly.
          )");

        PyFuture<PyObject*>::bind(mod, "_PyObjectFuture");

        py::class_<ExcelProcessInfo>(mod, "ExcelState", 
          R"(
            Gives information about the Excel application. Cannot be constructed: call
            ``xloil.excel_state`` to get an instance.
          )")
          .def_readonly("version", 
            &ExcelProcessInfo::version, 
            "Excel major version")
          .def_property_readonly("hinstance",
            [](const ExcelProcessInfo& p) { return (intptr_t)p.hInstance; },
            "Excel Win32 HINSTANCE pointer as an int")
          .def_readonly("hwnd",
            &ExcelProcessInfo::hWnd,
            "Excel Win32 main window handle as an int")
          .def_readonly("main_thread_id",
            &ExcelProcessInfo::mainThreadId,
            "Excel main thread ID");

        mod.def("excel_state", 
          Environment::excelProcess, 
          R"(
            Gives information about the Excel application, in particular the handles required
            to interact with Excel via the Win32 API. Only available when xlOil is loaded as 
            an addin.
          )",
          py::return_value_policy::reference);

        {
          auto e = py::exception<CannotConvert>(mod, "CannotConvert");
          e.doc() = R"(
            Should be thrown by a converter when it is unable to handle the 
            provided type.  In a return converter it may not indicate a fatal 
            condition, as xlOil will fallback to another converter.
          )";
          cannotConvertException = e.ptr();
        }

        // TODO: move to basictypes but beware of pybind declaration order!
        {
          // Bind CellError type to xloil::CellError enum
          auto eType = py::BetterEnum<CellError>(mod, "CellError", 0x800a07D0,
            R"(
              Enum-type class which represents an Excel error condition of the 
              form `#N/A!`, `#NAME!`, etc passed as a function argument. If a 
              registered function argument does not explicitly specify a type 
              (e.g. int or str via an annotation), it may be passed a *CellError*, 
              which it can handle based on the error type.

              The integer value of a *CellError* corresponds to it's VBA/COM error
              number, so for example we can write 
              `if cell.Value2 == CellError.NA.value: ...`
            )");
          
          for (auto e : theCellErrors)
            eType.value(cellErrorSymbol(e).c_str(), e);

          theCellErrorType = (PyTypeObject*)eType.ptr();
        }
      }
    }
} }