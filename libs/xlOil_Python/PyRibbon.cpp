#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include "EventLoop.h"
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
        
        return [pyMapper = PyObjectHolder(pyMapper), addin = &theCurrentAddin()](const wchar_t* name)
        {
          try
          {  
            return [&, funcname=wstring(name)](
              const RibbonControl& ctrl, VARIANT* vRet, int nArgs, VARIANT** vArgs)
              {
                try
                {
                  // Converting via an ExcelObj is not optimal, but avoids building
                  // a VARIANT converter. Since very few VARIANT types can appear in
                  // callbacks this might be a feasible approach when IPictureDisp
                  // support is introduced.
                  py::gil_scoped_acquire getGil;
                  auto callback = pyMapper(funcname);

                  py::tuple args(nArgs);
                  for (auto i = 0; i < nArgs; ++i)
                    args[i] = PyFromAny()(variantToExcelObj(*vArgs[i]));
                  
                  auto pyRet = callback(ctrl, *args);
                  auto isAsync = py::module::import("inspect").attr("iscoroutine")(pyRet).cast<bool>();

                  if (isAsync)
                  {
                    if (vRet)
                      XLO_THROW("Ribbon callback functions which return a value cannot be async");
                    addin->thread->runAsync(pyRet);
                  }
                  else if (vRet && !pyRet.is_none())
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

      auto setRibbon(IComAddin& addin, wstring xml, py::object mapper)
      {
        auto fut = runExcelThread([&addin, xml, maps = makeRibbonNameMapper(mapper)]()
        {
          addin.setRibbon(xml.c_str(), maps);
        });
        py::gil_scoped_release release;
        return fut.get();
      }

      inline auto makeAddin(wstring&& name)
      {
        auto fut = runExcelThread([name]()
        {
          return makeComAddin(name.c_str(), nullptr);
        });

        py::gil_scoped_release release;
        return fut.get();
      }

      inline auto makeAddin(
        wstring&& name,
        wstring&& xml,
        IComAddin::RibbonMap&& mapper)
      {
        auto fut = runExcelThread([name, xml, mapper]()
        {
          auto addin =  makeComAddin(name.c_str(), nullptr);
          addin->setRibbon(xml.c_str(), mapper);
          addin->connect();
          return addin;
        });

        py::gil_scoped_release release;
        return fut.get();
      }

      shared_ptr<IComAddin> createRibbon(const py::object& xml, const py::object& funcNameMap, const py::object& name)
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
          return makeAddin(std::move(addinName), std::move(xmlStr), std::move(mapper));
        }
        else
          return makeAddin(std::move(addinName));
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
      auto createTaskPane(IComAddin& addin, const std::wstring& name,
        const py::object& progId, const py::object& window)
      {
        auto winPtr = window.is_none() ? nullptr : (IDispatch*)ExcelWindow(pyToWStr(window).c_str()).ptr();
        auto progIdStr = progId.is_none() ? wstring() : pyToWStr(progId).c_str();

        auto fut = runExcelThread([&addin, name, winPtr, progIdStr]()
        {
          addin.connect();
          return addin.createTaskPane(name.c_str(), winPtr, progIdStr.empty() ? nullptr : progIdStr.c_str());
        });

        py::gil_scoped_release releaseGil;
        return fut.get();
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
        runExcelThread([&self, handler = make_shared<PyTaskPaneHandler>(eventHandler)]() 
        {
          self.addEventHandler(handler);
        });
      }
 
      void setTaskPaneSize(ICustomTaskPane* pane, const py::tuple& pair)
      {
        runExcelThread([pane, w = pair[0].cast<int>(), h = pair[1].cast<int>()]
        {
          pane->setSize(w, h);
        });
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<RibbonControl>(mod, "RibbonControl")
          .def_readonly("id", &RibbonControl::Id)
          .def_readonly("tag", &RibbonControl::Tag);

        py::class_<ICustomTaskPane, shared_ptr<ICustomTaskPane>>(mod, "TaskPaneFrame")
          .def_property_readonly("parent_hwnd", 
            &ICustomTaskPane::parentWindowHandle)
          .def_property_readonly("window", 
            MainThreadWrap(&ICustomTaskPane::window))
          .def_property("visible", 
            MainThreadWrap(&ICustomTaskPane::getVisible), MainThreadWrap(&ICustomTaskPane::setVisible))
          .def_property("size", 
            MainThreadWrap(&ICustomTaskPane::getSize), setTaskPaneSize)
          .def_property_readonly("title", 
            MainThreadWrap(&ICustomTaskPane::getTitle))
          .def("com_control", 
            &ICustomTaskPane::content)
          .def("add_event_handler", 
            &addPaneEventHandler, py::arg("handler"));

        py::class_<IComAddin, shared_ptr<IComAddin>>(mod, "ExcelUI")
          .def(py::init(std::function(createRibbon)), py::arg("ribbon")=py::none(), py::arg("func_names")=py::none(), py::arg("name")=py::none())
          .def("connect", 
            MainThreadWrap(&IComAddin::connect))
          .def("disconnect", 
            MainThreadWrap(&IComAddin::disconnect))
          .def("ribbon", 
            setRibbon, py::arg("xml"), py::arg("func_names"))
          .def("invalidate", 
            MainThreadWrap([](IComAddin* p, wstring id) { return p->ribbonInvalidate(id.c_str()); }), py::arg("id") = "")
          .def("activate", 
            MainThreadWrap([](IComAddin* p, wstring id) { return p->ribbonActivate(id.c_str()); }), py::arg("id"))
          .def("add_task_pane", 
            createTaskPane, py::arg("name"), py::arg("progid")=py::none(), py::arg("window")=py::none())
          .def_property_readonly("name", 
            &IComAddin::progid);
      });
    }
  }
}