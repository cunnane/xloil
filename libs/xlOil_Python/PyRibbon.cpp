#include "Main.h"
#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include "EventLoop.h"
#include "PyFuture.h"
#include "PyCom.h"
#include <xloil/ExcelUI.h>
#include <xlOil/AppObjects.h>
#include <xloil/ExcelThread.h>
#include <xloil/RtdServer.h>
#include <pybind11/pybind11.h>
#include <filesystem>
#include <OleAuto.h>

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

      using AddinFuture = PyFuture<shared_ptr<IComAddin>>;
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

      using CTPFuture = PyFuture<shared_ptr<ICustomTaskPane>>;

      CTPFuture createPaneFrame(
        IComAddin& comAddin, 
        const std::wstring& name,
        const py::object& window,
        const py::object& progId)
      {
        auto progIdStr = progId.is_none() ? wstring() : pyToWStr(progId).c_str();
        auto winPtr    = window.is_none() ? nullptr   : window.cast<ExcelWindow>().dispatchPtr();

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
        return py::module::import("xloil").attr("create_task_pane")(*args, **kwargs);
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

        py::class_<RibbonControl>(mod, 
          "RibbonControl", R"(
            This object is passed to ribbon callback handlers to indicate which control  
            raised the callback.
          )")
          .def_readonly("id", 
            &RibbonControl::Id,
            "A string that represents the Id attribute for the control or custom menu item")
          .def_readonly("tag", 
            &RibbonControl::Tag,
            "A string that represents the Tag attribute for the control or custom menu item.");

        py::class_<ICustomTaskPane, shared_ptr<ICustomTaskPane>>(mod, 
          "TaskPaneFrame", R"(
            References Excel's base task pane object into which the python GUI can be drawn.
            The methods of this object are safe to call from any thread.  COM must be used on Excel's
            main thread, so the methods all wrap their calls to ensure to this happens. This could lead 
            to deadlocks if the call triggers event  handlers on the main thread, which in turn block 
            waiting for the thread originally calling `TaskPaneFrame`.
          )")
          .def_property_readonly("parent_hwnd", 
            &ICustomTaskPane::parentWindowHandle,
            "Win32 window handle used to attach a python GUI to a task pane frame")
          .def_property_readonly("window", 
            MainThreadWrap(&ICustomTaskPane::window),
            "Gives the window of the document window to which the frame is attached, can be "
            "used to uniquely identify the pane")
          .def_property("visible", 
            MainThreadWrap(&ICustomTaskPane::getVisible), 
            MainThreadWrap(&ICustomTaskPane::setVisible),
            "Determines the visibility of the task pane")
          .def_property("size", 
            MainThreadWrap(&ICustomTaskPane::getSize), 
            setTaskPaneSize,
            "Gets/sets the task pane size as a tuple (width, height)")
          .def_property_readonly("title", 
            MainThreadWrap(&ICustomTaskPane::getTitle))
          .def("com_control", 
            [](ICustomTaskPane& self, const char* binder)
            {
              py::gil_scoped_release noGil;
              return comToPy(*self.content(), binder);
            },
            R"(
              Gets the base COM control of the task pane. The ``lib`` used to provide
              COM support can be 'comtypes' or 'win32com' (default is win32com).
            )",
            py::arg("lib") = "")
          .def("add_event_handler", 
            &addPaneEventHandler, py::arg("handler"));

        py::class_<IComAddin, shared_ptr<IComAddin>>(mod, 
          "ExcelGUI", R"(
            Controls an Ribbon and its associated COM addin. The methods of this object are safe
            to call from any thread.  However, COM must be used on Excel's main thread, so the methods  
            schedule calls to run on the main thead. This could lead to deadlocks if the call 
            triggers event handlers on the main thread, which in turn block whilst waiting for the 
            thread originally calling ExcelGUI.
          )")
          .def("connect",
            comAddin_connect, 
            R"(
              Connects this COM add-in underlying this Ribbon to Excel. Any specified 
              ribbon XML will be passed to Excel.
            )",
            py::arg("xml")="", 
            py::arg("func_names")=py::none())
          .def("disconnect",
            MainThreadWrap(&IComAddin::disconnect),
            "Unloads the underlying COM add-in and any ribbon customisation.")
          .def("invalidate",
            MainThreadWrap([](IComAddin* p, const wstring& id) { return p->ribbonInvalidate(id.c_str()); }),
            R"(
              Invalidates the specified control: this clears the cache of responses
              to callbacks associated with the control. For example, this can be
              used to hide a control by forcing its getVisible callback to be invoked,
              rather than using the cached value.

              If no control ID is specified, all controls are invalidated.
            )",
            py::arg("id") = "")
          .def("activate",
            MainThreadWrap([](IComAddin* p, const wstring& id) { return p->ribbonActivate(id.c_str()); }),
            R"(
              Activatives the ribbon tab with the specified id.  Returns False if
              there is no Ribbon or the Ribbon is collapsed.
            )",
            py::arg("id"))
          .def("task_pane_frame",
            createPaneFrame, 
            R"(
              Used internally to create a custom task pane window which can be populated
              with a python GUI.  Most users should use `create_task_pane(...)` instead.

              A COM `progid` can be specified, but this will prevent using a python GUI
              in the task pane. This is a specialised use case.
            )",
            py::arg("name"), 
            py::arg("progid") = py::none(), 
            py::arg("window") = py::none())
          .def("create_task_pane", 
            createTaskPane,
            R"(
              Returns a task pane with title <name> attached to the active window,
              creating it if it does not already exist.  See `xloil.create_task_pane`.

              Parameters
              ----------

              creator: 
                  * a subclass of `QWidget` or
                  * a function which takes a `TaskPaneFrame` and returns a `CustomTaskPane`

              window: 
                  a window title or `ExcelWindow` object to which the task pane should be
                  attached.  If None, the active window is used.
            )")
          .def_property_readonly("name", 
            &IComAddin::progid);

        mod.def("create_gui", 
          createRibbon, 
          R"(
            Returns an **awaitable** to a ExcelGUI object which passes the specified ribbon
            customisation XML to Excel.  When the returned object is deleted, it 
            unloads the Ribbon customisation and the associated COM add-in.  If ribbon
            XML is specfied the ExcelGUI object will be connected, otherwise the 
            user must call the `connect()` method to active the object.

            Parameters
            ----------

            ribbon: str
                A Ribbon XML string, most easily created with a specialised editor.
                The XML format is documented on Microsoft's website

            func_names: Func[str -> callable] or Dict[str, callabe]
                The ``func_names`` mapper links callbacks named in the Ribbon XML to
                python functions. It can be either a dictionary containing named 
                functions or any callable which returns a function given a string.
                Each return handler should take a single ``RibbonControl``
                argument which describes the control which raised the callback.

                Callbacks declared async will be executed in the addin's event loop. 
                Other callbacks are executed in Excel's main thread. Async callbacks 
                cannot return values.

            name: str
                The addin name which will appear in Excel's COM addin list.
                If None, uses the filename at the call site as the addin name.
          )",
          py::arg("ribbon") = py::none(), 
          py::arg("func_names") = py::none(), 
          py::arg("name") = py::none());
      });
    }
  }
}