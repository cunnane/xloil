#include "PyCoreModule.h"
#include "PyHelpers.h"
#include "PyExcelArray.h"
#include "BasicTypes.h"
#include <xlOil/ExcelApp.h>
#include <xloil/Log.h>
#include <xloil/Caller.h>
#include <xloil/State.h>
#include <map>

using std::shared_ptr;
using std::vector;
using std::wstring;
namespace py = pybind11;
using std::make_pair;

namespace xloil {
  namespace Python {

    using BinderFunc = std::function<void(pybind11::module&)>;
    void bindFirst(py::module& mod);
    namespace
    {
      class BinderRegistry
      {
      public:
        static BinderRegistry& get() {
          static BinderRegistry instance;
          return instance;
        }

        void add(BinderFunc f, size_t priority)
        {
          theFunctions.insert(make_pair(priority, f));
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
      auto mod = py::module(theInjectedModuleName);
      BinderRegistry::get().bindAll(mod);
      return mod.release().ptr();
    }

    int addBinder(std::function<void(pybind11::module&)> binder, size_t priority)
    {
      BinderRegistry::get().add(binder, priority);
      return 0;
    }

    namespace
    {
      void writeToLog(const char* message, const char* level)
      {
        const auto levelEnum = spdlog::level::from_str(level);
        if (levelEnum != spdlog::level::off)
          SPDLOG_LOGGER_CALL(spdlog::default_logger_raw(), levelEnum, message);
      }

      void runLater(const py::object& callable, int nRetries, int retryPause, int delay)
      {
        excelPost([callable]()
          {
            py::gil_scoped_acquire getGil;
            callable();
          },
          QueueType::WINDOW,
          nRetries,
          retryPause,
          delay);
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        // Bind the two base classes for python converters
        py::class_<IPyFromExcel, shared_ptr<IPyFromExcel>>(mod, "IPyFromExcel")
          .def("__call__",
            [](const IPyFromExcel& /*self*/, const py::object& /*arg*/)
            {
              XLO_THROW("Not implemented");
            });
        py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

        mod.def("in_wizard", &inFunctionWizard);
        mod.def("log", &writeToLog, py::arg("msg"), py::arg("level") = "info");
        mod.def("run_later",
          &runLater,
          py::arg("func"),
          py::arg("num_retries") = 10,
          py::arg("retry_delay") = 500,
          py::arg("wait_time") = 0);

        py::class_<State::ExcelState>(mod, "ExcelState")
          .def_readonly("version", &State::ExcelState::version)
          .def_readonly("hinstance", &State::ExcelState::hInstance)
          .def_readonly("hwnd", &State::ExcelState::hWnd)
          .def_readonly("main_thread_id", &State::ExcelState::mainThreadId);

        mod.def("get_excel_state", State::excelState);

        py::class_<CallerInfo>(mod, "Caller")
          .def(py::init<>())
          .def_property_readonly("sheet",
            [](const CallerInfo& self) 
            { 
              const auto name = self.sheetName();
              return name.empty() ? py::none() : py::wstr(wstring(name));
            })
          .def_property_readonly("workbook",
            [](const CallerInfo& self)
            {
              const auto name = self.workbook();
              return name.empty() ? py::none() : py::wstr(wstring(name));
            })
          .def("address", [](const CallerInfo& self, bool x)
            {
              return self.writeAddress(x);
            }, py::arg("a1style") = false);

      }, 1000);
    }
} }