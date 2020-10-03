#include "Main.h"
#include "BasicTypes.h"
#include "PyCoreModule.h"
#include <xloil/Ribbon.h>
#include <xloil/RtdServer.h>
#include <pybind11/pybind11.h>
namespace py = pybind11;
using std::shared_ptr;
using std::wstring;

namespace xloil
{
  namespace Python
  {
    auto setRibbon(IComAddin* addin, const wchar_t* xml, const py::dict& handlers)
    {
      std::map<std::wstring, std::function<void(const RibbonControl&)>> funcs;
      for (auto[name, fn] : handlers)
      {
        funcs[py::cast<std::wstring>(name)] =
          [callable = PyBorrow<py::function>(fn.ptr())](const RibbonControl&)
        {
          py::gil_scoped_acquire getGil;
          callable.call();
        };
      }
      addin->setRibbon(xml, funcs);
    }
    auto createRibbon(const wchar_t* xml, const py::dict& handlers)
    {
      auto addin = makeComAddin(theCurrentContext->fileName());
      setRibbon(addin.get(), xml, handlers);
      addin->connect();
      return addin;
    }
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<RibbonControl>(mod, "RibbonControl")
          .def_readonly("id", &RibbonControl::Id)
          .def_readonly("tag", &RibbonControl::Tag);
        py::class_<IComAddin, shared_ptr<IComAddin>>(mod, "RibbonUI")
          .def("connect", &IComAddin::connect)
          .def("disconnect", &IComAddin::disconnect)
          .def("set_ribbon", setRibbon, py::arg("xml"), py::arg("handlers"))
          .def("invalidate", &IComAddin::ribbonInvalidate, py::arg("id") = nullptr)
          .def("activate", &IComAddin::ribbonActivate, py::arg("id"));
        mod.def("create_ribbon", createRibbon, py::arg("xml"), py::arg("handlers"));
      });
    }
  }
}