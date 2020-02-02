#include "InjectedModule.h"
#include "PyHelpers.h"
#include "xloil/Log.h"

using std::shared_ptr;
using std::vector;
namespace py = pybind11;

namespace xloil {
  namespace Python {

    using BinderFunc = std::function<void(pybind11::module&)>;
    namespace
    {
      class BinderRegistry
      {
      public:
        static BinderRegistry& get() {
          static BinderRegistry instance;
          return instance;
        }

        void add(BinderFunc f)
        {
          theFunctions.push_back(f);
        }

        void bindAll(py::module& mod)
        {
          for (auto m : theFunctions)
            m(mod);
        }
      private:
        BinderRegistry() {}
        vector<BinderFunc> theFunctions;
      };
    }

    PyObject* buildInjectedModule()
    {
      auto mod = py::module(XLO_PY_MOD_STR);
      BinderRegistry::get().bindAll(mod);
      return mod.release().ptr();
    }

    int addBinder(std::function<void(pybind11::module&)> binder)
    {
      BinderRegistry::get().add(binder);
      return 0;
    }

    void scanModule(py::object& mod)
    {
      py::gil_scoped_acquire get_gil;

      auto oilModule = py::module::import("xloil");
      auto scanFunc = oilModule.attr("scan_module").cast<py::function>();
            
      try
      {
        XLO_INFO("Scanning module {0}", (std::string)py::str(mod));
        scanFunc.call(mod);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Error reading module {0}: {1}", (std::string)py::str(mod) , e.what());
      }
    }
} }