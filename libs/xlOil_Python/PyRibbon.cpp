#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include <xloil/ExcelUI.h>
#include <xloil/RtdServer.h>
#include <pybind11/pybind11.h>
#include <filesystem>
namespace py = pybind11;
using std::shared_ptr;
using std::unique_ptr;
using std::wstring;
using std::vector;
using std::make_shared;
using std::make_unique;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      /// <summary>
      /// Expects funcNameMap to be either a function of name -> handler or a 
      /// dict of names and handler.  The handler function arguments vary depending
      /// on what type of ribbon callback is requested
      /// </summary>
      auto makeRibbonNameMapper(const py::object& funcNameMap)
      {
        auto pyMapper = PyDict_Check(funcNameMap.ptr())
          ? funcNameMap.attr("__getitem__")
          : funcNameMap;

        return [pyMapper = PyObjectHolder(pyMapper)](const wchar_t* name)
        {
          try
          {
            py::gil_scoped_acquire getGil;
            auto callback = pyMapper(name);
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
                auto pyRet = callback(ctrl, *args);
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
                    excelObjToVariant(vRet, FromPyObj<false>()(pyRet.ptr()));
                }
              }
              catch (const py::error_already_set& e)
              {
                raiseUserException(e);
                throw;
              }
            };
          }
          catch (const py::error_already_set& e)
          {
            raiseUserException(e);
            throw;
          }
        };
      }

      auto setRibbon(IComAddin& addin, const wchar_t* xml, py::object mapper)
      {
        addin.setRibbon(xml, makeRibbonNameMapper(mapper));
      }

      auto createRibbon(const py::object& xml, const py::object& funcNameMap, const py::object& name)
      {
        wstring addinName;
        if (name.is_none())
        {
          // The returned pointers here do no need to be freed or decref'd
          auto frame = PyEval_GetFrame();
          if (!frame)
            throw py::cast_error();
  #if PY_MAJOR_VERSION >= 4 || PY_MINOR_VERSION >= 9       
          auto code = PyFrame_GetCode(frame);
  #else
          auto code = frame->f_code;
  #endif
          std::filesystem::path filePath(pyToWStr(code->co_filename));
          addinName = filePath.filename().stem();
        }
        else
          addinName = pyToWStr(name.ptr());
        
        if (!xml.is_none())
        {
          auto mapper = makeRibbonNameMapper(funcNameMap);
          auto xmlStr = pyToWStr(xml);

          py::gil_scoped_release releaseGil;
          auto addin = make_unique<ComAddinThreadSafe>(
            addinName.c_str(), xmlStr.c_str(), std::move(mapper));
          return addin.release();
        }
        else
        {
          auto addin = make_unique<ComAddinThreadSafe>(addinName.c_str());
          return addin.release();
        }
      }
      
      PyObject* callOneArg(PyObject* callable, PyObject* arg)
      {
#if PY_VERSION_HEX < 0x03080000
        auto result = PyObject_CallFunctionObjArgs(callable, arg, nullptr);
#elseif PY_VERSION_HEX < 0x03090000
          PyObject * args[] = { nullptr, arg };
        auto result = _PyObject_Vectorcall(callable, args + 1, 1 | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#else
        auto result = PyObject_CallOneArg(callable, arg);
#endif

        return result;
      }

      // TODO: attach task pane to any windowCaption
      auto createTaskPane(IComAddin& self, const std::wstring& name, 
        const py::object& progId, const py::object& window)
      {
        auto winPtr = window.is_none() ? nullptr : (IDispatch*)ExcelWindow(pyToWStr(window).c_str()).ptr();
        auto progIdStr = progId.is_none() ? wstring() : pyToWStr(progId).c_str();

        py::gil_scoped_release releaseGil;
        self.connect();
        return self.createTaskPane(name.c_str(), winPtr, progIdStr.empty() ? nullptr : progIdStr.c_str());
      }

      class PyTaskPaneHandler : public ICustomTaskPaneHandler
      {
      public:
        PyTaskPaneHandler(const py::object& eventHandler)
          : _handler(eventHandler)
        {}

        void onSize(int width, int height) override
        {
          py::gil_scoped_acquire gil;
          checkUserException([=]() { _handler.attr("on_size")(width, height); });
        }
        void onVisible(bool c) override
        {
          py::gil_scoped_acquire gil;
          checkUserException([=]() { _handler.attr("on_visible")(c); });
        }
        void onDocked() override
        {
          py::gil_scoped_acquire gil;
          checkUserException([this]() { _handler.attr("on_docked")(); });
        }
        void onDestroy() override
        {
          py::gil_scoped_acquire gil;
          checkUserException([this]() { _handler.attr("on_destroy")(); });
        }
        PyObjectHolder _handler;
      };

      void addPaneEventHandler(ICustomTaskPane& self, const py::object& eventHandler)
      {
        self.addEventHandler(make_shared<PyTaskPaneHandler>(eventHandler));
      }
 
      void setTaskPaneSize(ICustomTaskPane* pane, const py::tuple& pair)
      {
        pane->setSize(pair[0].cast<int>(), pair[1].cast<int>());
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<RibbonControl>(mod, "RibbonControl")
          .def_readonly("id", &RibbonControl::Id)
          .def_readonly("tag", &RibbonControl::Tag);

        py::class_<ICustomTaskPane, shared_ptr<ICustomTaskPane>>(mod, "TaskPaneFrame")
          .def_property_readonly("parent_hwnd", &ICustomTaskPane::parentWindowHandle)
          .def_property_readonly("window", &ICustomTaskPane::window)
          .def_property("visible", &ICustomTaskPane::getVisible, &ICustomTaskPane::setVisible)
          .def_property("size", &ICustomTaskPane::getSize, setTaskPaneSize)
          .def_property_readonly("title", &ICustomTaskPane::getTitle)
          .def("com_control", &ICustomTaskPane::content)
          .def("add_event_handler", &addPaneEventHandler, py::arg("handler"));

        py::class_<IComAddin>(mod, "ExcelUI")
          .def(py::init(std::function(createRibbon)), py::arg("ribbon")=py::none(), py::arg("func_names")=py::none(), py::arg("name")=py::none())
          .def("connect", &IComAddin::connect)
          .def("disconnect", &IComAddin::disconnect)
          .def("ribbon", setRibbon, py::arg("xml"), py::arg("func_names"))
          .def("invalidate", &IComAddin::ribbonInvalidate, py::arg("id") = nullptr)
          .def("activate", &IComAddin::ribbonActivate, py::arg("id"))
          .def("add_task_pane", createTaskPane, py::arg("name"), py::arg("progid")=py::none(), py::arg("window")=py::none())
          .def_property_readonly("name", &IComAddin::progid);
      });
    }
  }
}