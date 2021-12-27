#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include "EventLoop.h"
#include "PyFuture.h"
#include <xloil/ExcelUI.h>
#include <xloil/ExcelThread.h>
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
        
        return [pyMapper = PyObjectHolder(pyMapper), eventLoop = getEventLoop()](
          const wchar_t* name)
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
                    eventLoop->runAsync(pyRet);
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

      using AddinFuture = PyFuture<shared_ptr<IComAddin>, detail::CastFutureConverter>;
      using VoidFuture = PyFuture<void, void>;

      inline auto comAddin_connect(IComAddin& addin, const wstring& xml, const py::object& funcmap)
      {
        return new VoidFuture(runExcelThread([
            &addin, 
            xml, 
            mapper = funcmap.is_none() ? IComAddin::RibbonMap() : makeRibbonNameMapper(funcmap)]
          () {
            if (!xml.empty())
              addin.setRibbon(xml.c_str(), mapper);
            addin.connect();
          }));
      }

      inline auto makeAddin(wstring&& name)
      {
        return new AddinFuture(runExcelThread([name]()
        {
          return makeComAddin(name.c_str(), nullptr);
        }));
      }

      inline auto makeAddin(
        wstring&& name,
        wstring&& xml,
        IComAddin::RibbonMap&& mapper)
      {
        return new AddinFuture(runExcelThread([name, xml, mapper]()
        {
          auto addin = makeComAddin(name.c_str(), nullptr);
          addin->setRibbon(xml.c_str(), mapper);
          addin->connect();
          return addin;
        }));
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
          return makeAddin(std::move(addinName), std::move(xmlStr), std::move(mapper));
        }
        else
          return makeAddin(std::move(addinName));
      }
      
      PyObject* callOneArg(PyObject* callable, PyObject* arg)
      {
#if PY_VERSION_HEX < 0x03080000
        auto result = PyObject_CallFunctionObjArgs(callable, arg, nullptr);
#elif PY_VERSION_HEX < 0x03090000
          PyObject * args[] = { nullptr, arg };
        auto result = _PyObject_Vectorcall(callable, args + 1, 1 | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#else
        auto result = PyObject_CallOneArg(callable, arg);
#endif

        return result;
      }

      using CTPFuture = PyFuture<shared_ptr<ICustomTaskPane>,detail::CastFutureConverter>;

      CTPFuture createPaneFrame(
        IComAddin& comAddin, 
        const std::wstring& name,
        const py::object& window,
        const py::object& progId)
      {
        auto progIdStr = progId.is_none() ? wstring() : pyToWStr(progId).c_str();
        auto winPtr    = window.is_none() ? nullptr   : window.cast<ExcelWindow>().basePtr();

        return runExcelThread([&comAddin, name, winPtr, progIdStr]()
        {
          comAddin.connect();
          return comAddin.createTaskPane(name.c_str(), winPtr, progIdStr.empty() ? nullptr : progIdStr.c_str());
        });
      }

      auto createTaskPane(
        IComAddin& comAddin, 
        py::args args, 
        py::kwargs kwargs)
      {
        kwargs["gui"] = py::cast(comAddin);
        return py::module::import("xloil.excelgui").attr("create_task_pane")(*args, **kwargs);
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
        CTPFuture::bind(mod, "_CTPFuture");
        AddinFuture::bind(mod, "_AddinFuture");
        VoidFuture::bind(mod);

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

        py::class_<IComAddin, shared_ptr<IComAddin>>(mod, "ExcelGUI")
          .def("connect",
            comAddin_connect, py::arg("xml")="", py::arg("func_names")=py::none())
          .def("disconnect",
            MainThreadWrap(&IComAddin::disconnect))
          .def("invalidate",
            MainThreadWrap([](IComAddin* p, const wstring& id) { return p->ribbonInvalidate(id.c_str()); }), py::arg("id") = "")
          .def("activate",
            MainThreadWrap([](IComAddin* p, const wstring& id) { return p->ribbonActivate(id.c_str()); }), py::arg("id"))
          .def("task_pane_frame",
            createPaneFrame, py::arg("name"), py::arg("progid") = py::none(), py::arg("window") = py::none())
          .def("create_task_pane", 
            createTaskPane)
          .def_property_readonly("name", 
            &IComAddin::progid);

        mod.def("create_gui", createRibbon, py::arg("ribbon") = py::none(), py::arg("func_names") = py::none(), py::arg("name") = py::none());
      });
    }
  }
}