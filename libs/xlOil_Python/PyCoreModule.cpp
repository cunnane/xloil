#include "PyCoreModule.h"
#include "PyHelpers.h"
#include "PyExcelArray.h"
#include "BasicTypes.h"
#include <xloil/ApiMessage.h>
#include <xloil/Log.h>
#include <map>

using std::shared_ptr;
using std::vector;
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
        SPDLOG_LOGGER_CALL(spdlog::default_logger_raw(), spdlog::level::from_str(level), message);
      }

      void runLater(const py::object& callable, int nRetries, int retryPause, int delay)
      {
        excelApiCall([callable]()
          {
            py::gil_scoped_acquire getGil;
            callable.call();
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
            [](const IPyFromExcel& self, const py::object& arg)
            {
              if (Py_TYPE(arg.ptr()) == ExcelArrayType)
              {
                auto arr = arg.cast<PyExcelArray>();
                return self.fromArray(arr.base());
              }
              else if (PyLong_Check(arg.ptr()))
              {
                return self(ExcelObj(arg.cast<long>()));
              }
              XLO_THROW("Not implemented");
            });
        py::class_<IPyToExcel, shared_ptr<IPyToExcel>>(mod, "IPyToExcel");

        mod.def("in_wizard", &Core::inFunctionWizard);
        mod.def("log", &writeToLog, py::arg("msg"), py::arg("level") = "info");
        mod.def("run_later",
          &runLater,
          py::arg("func"),
          py::arg("num_retries") = 10,
          py::arg("retry_delay") = 500,
          py::arg("wait_time") = 0);
      }, 1000);
    }
} }