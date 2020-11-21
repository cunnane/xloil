#include "Main.h"
#include "BasicTypes.h"
#include "PyCoreModule.h"
#include "PyEvents.h"
#include "PyImage.h"
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
    auto setRibbon(IComAddin* addin, const wchar_t* xml, py::object mapper)
    {
      if (PyDict_Check(mapper.ptr()))
        mapper = mapper.attr("__getitem__");

      auto cmapper = [mapper](const wchar_t* name)
      {
        try
        {
          py::gil_scoped_acquire getGil;
          auto callback = mapper.call(name);
          return [callback](
            const RibbonControl& ctrl, VARIANT* vRet, int nArgs, VARIANT** vArgs)
          {
            try
            {
              // Converting via an ExcelObj is not optimal, but avoids building
              // a VARIANT converter. Since very few VARIANT types can appear in
              // callbacks this might be a feasible approach when IPictureDisp
              // support is introduced.
              py::gil_scoped_acquire getGil;
              py::tuple args(nArgs);
              for (auto i = 0; i < nArgs; ++i)
                args[i] = PyFromAny()(variantToExcelObj(*vArgs[i]));
              auto pyRet = callback.call(ctrl, *args);
              if (vRet && !pyRet.is_none())
              {
                auto picture = pictureFromPilImage(pyRet);
                if (picture)
                {
                  VariantInit(vRet);
                  vRet->pdispVal = (IDispatch*)picture;
                  vRet->vt = VT_DISPATCH;
                }
                else
                  *vRet = excelObjToVariant(FromPyObj()(pyRet.ptr(), false));
              }
            }
            catch (const py::error_already_set& e)
            {
              Event_PyUserException().fire(e.type(), e.value(), e.trace());
              throw;
            }
          };
        }
        catch (const py::error_already_set& e)
        {
          Event_PyUserException().fire(e.type(), e.value(), e.trace());
          throw;
        }
      };
      addin->setRibbon(xml, cmapper);
    }
    auto createRibbon(const wchar_t* xml, const py::object& mapper)
    {
      auto addin = makeComAddin(theCurrentContext->fileName());
      setRibbon(addin.get(), xml, mapper);
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
          .def("set_ribbon", setRibbon, py::arg("xml"), py::arg("mapper"))
          .def("invalidate", &IComAddin::ribbonInvalidate, py::arg("id") = nullptr)
          .def("activate", &IComAddin::ribbonActivate, py::arg("id"));
        mod.def("create_ribbon", createRibbon, py::arg("xml"), py::arg("mapper"));
      });
    }
  }
}