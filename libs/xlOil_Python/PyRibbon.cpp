#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyCore.h"
#include "PyEvents.h"
#include "PyImage.h"
#include "EventLoop.h"
#include "PyFuture.h"
#include "PyCom.h"
#include "PyAddin.h"

#include <xloil/ExcelUI.h>
#include <xlOil/AppObjects.h>
#include <xloil/ExcelThread.h>
#include <xloil/RtdServer.h>
#include <pybind11/pybind11.h>
#include <OleAuto.h>
#include <filesystem>

namespace py = pybind11;
using std::shared_ptr;
using std::unique_ptr;
using std::wstring;
using std::vector;
using std::make_shared;
using std::make_unique;
using std::pair;

namespace xloil
{
  namespace Python
  {
    namespace
    {
      struct PyRibbonControl
      {
        PyRibbonControl(const RibbonControl& ctrl)
        {
          if (ctrl.Id)  Id  = ctrl.Id;
          if (ctrl.Tag) Tag = ctrl.Tag;
        }
        wstring Id;
        wstring Tag;
      };

      constexpr static pair<ICustomTaskPane::DockPosition, const char*> positionNames[] = {
        pair(ICustomTaskPane::Bottom, "bottom"),
        pair(ICustomTaskPane::Floating, "floating"),
        pair(ICustomTaskPane::Left, "left"),
        pair(ICustomTaskPane::Right, "right"),
        pair(ICustomTaskPane::Top, "top")
      };

      const char* TaskPane_getPosition(ICustomTaskPane& self)
      {
        py::gil_scoped_release noGil;
        auto position = runExcelThread([&]() {
          return self.getPosition();
        }).get();
        for (auto posName : positionNames)
        {
          if (position == posName.first)
            return posName.second;
        }
        return "";
      }

      void TaskPane_setPosition(ICustomTaskPane& self, std::string position)
      {
#pragma warning(disable: 4244)
        std::transform(position.begin(), position.end(), position.begin(), ::tolower);
        for (auto posName : positionNames)
        {
          if (position == posName.second)
          {
            runExcelThread([&]() {
              self.setPosition(posName.first);
            });
            return;
          }
        }
        throw new py::value_error("Unrecognised position: '" + position + "'");
      }
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
        
        return [pyMapper = PyObjectHolder(pyMapper), eventLoop = getEventLoop().get()](
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
                  
                  auto pyRet = callback(PyRibbonControl(ctrl), *args);
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

      using VoidFuture = PyFuture<void, void>;
      using CTPFuture = PyFuture<shared_ptr<ICustomTaskPane>>;

      class ComAddin
      {
      public:
        ComAddin(
          const py::object& name, 
          const py::object& xml, 
          const py::object& funcmap,
          bool connect)
        {
          if (name.is_none())
          {
            // The returned pointers here do no need to be freed or decref'd
            auto frame = PyEval_GetFrame();
            if (!frame)
              throw py::cast_error();
#if PY_VERSION_HEX >= 0x03090000
            auto code = PyFrame_GetCode(frame);
#else
            auto code = frame->f_code;
#endif
            std::filesystem::path filePath(to_wstring(code->co_filename));
            _name = filePath.filename().stem();
          }
          else
            _name = to_wstring(name);

          if (!xml.is_none())
          {
            _xml = to_wstring(xml);
            _functionMap = funcmap;
          }
          if (connect)
          {
            this->connect().result();
          }
        }

        ~ComAddin()
        {
          py::gil_scoped_release noGil;
          runExcelThread([this]() { _addin.reset(); }).get();
        }

        VoidFuture connect()
        {
          auto ribbonMap = _functionMap.is_none()
            ? IComAddin::RibbonMap()
            : makeRibbonNameMapper(_functionMap);
          _functionMap = py::object();

          return runExcelThread([this, mapper = std::move(ribbonMap)]() mutable
          {
            if (!_addin)
              _addin = makeComAddin(_name.c_str(), nullptr);
            _addin->connect(_xml.c_str(), mapper);
            _connected = true;
          });
        }

        VoidFuture disconnect()
        {
          if (!_connected || !_addin)
          {
             std::promise<void> p; 
             p.set_value(); 
             return p.get_future();
          }
          return runExcelThread([this]() mutable 
          {
            _addin->disconnect(); 
            _connected = false;
          });
        }

        VoidFuture invalidate(const wstring& id)
        {
          return runExcelThread([addin = addin(), id]() { addin->ribbonInvalidate(id.c_str()); });
        }

        VoidFuture activate(const wstring& id)
        {
          return runExcelThread([addin = addin(), id]() { addin->ribbonActivate(id.c_str()); });
        }

        CTPFuture createTaskPaneFrame(
          const std::wstring& name,
          const py::object& window,
          const py::object& progId)
        {
          auto progIdStr = progId.is_none() ? wstring() : to_wstring(progId).c_str();
          auto winPtr = window.is_none() ? ExcelWindow(nullptr) : window.cast<ExcelWindow>();

          return runExcelThread([addin = addin(), name, winPtr, progIdStr]()
          {
            return addin->createTaskPane(
              name.c_str(), 
              &winPtr, 
              progIdStr.empty() ? nullptr : progIdStr.c_str());
          });
        }

        IComAddin* addin()
        {
          if (!_addin || !_connected) XLO_THROW("Addin disconnected");
          return _addin.get();
        }

        auto name() const { return _name; }

        bool connected() const { return _connected; }

      private:
        shared_ptr<IComAddin> _addin;
        wstring _name;
        wstring _xml;
        py::object _functionMap;
        std::atomic<bool> _connected;
      };
    
      auto attachTaskPaneAsync(
        const py::object& comAddin,
        const py::object& pane,
        const py::object& name,
        const py::object& window,
        const py::object& size,
        const py::object& visible)
      {
        auto attachPane = py::module::import("xloil.gui").attr("_attach_task_pane_async");
        return attachPane(comAddin, pane, name, window, size, visible);
      }

      auto attachTaskPane(
        const py::object& comAddin,
        const py::object& pane,
        const py::object& name,
        const py::object& window,
        const py::object& size,
        const py::object& visible)
      {
        auto attachPane = py::module::import("xloil.gui").attr("_attach_task_pane");
        return attachPane(comAddin, pane, name, window, size, visible);
      }
      
      auto createTaskPane(
        const py::object& comAddin,
        const py::object& name,
        const py::object& pane,
        const py::object& window,
        const py::object& size,
        const py::object& visible)
      {
        PyErr_WarnEx(PyExc_DeprecationWarning,
          "createTaskPane is deprecated, use attach_pane instead.",
          2);

        auto guiModule = py::module::import("xloil.gui");

        auto findPane = guiModule.attr("find_task_pane");
        auto found = findPane(name);
        if (!found.is_none())
          return found;

        return attachTaskPane(comAddin, pane, name, window, size, visible);
      }

      class PyTaskPaneHandler : public ICustomTaskPaneEvents
      {
      public:
        PyTaskPaneHandler(const py::object& eventHandler)
          : _handler(eventHandler)
        {
          _hasOnVisible = py::hasattr(eventHandler, "on_visible");
          _hasOnDocked  = py::hasattr(eventHandler, "on_docked");
          _hasOnDestroy = py::hasattr(eventHandler, "on_destroy");
        }
        ~PyTaskPaneHandler()
        {
          py::gil_scoped_acquire gil;
          _handler.dec_ref();
        }
        void onVisible(bool c) override
        {
          if (!_hasOnVisible) return;
          py::gil_scoped_acquire gil;
          auto handler = PyBorrow(_handler);
          if (handler.is_none()) return;
          checkUserException([=]() { handler.attr("on_visible")(c); });
        }
        void onDocked() override
        {
          if (!_hasOnDocked) return;
          py::gil_scoped_acquire gil;
          auto handler = PyBorrow(_handler);
          if (handler.is_none()) return;
          checkUserException([=]() { handler.attr("on_docked")(); });
        }
        void onDestroy() override
        {
          if (!_hasOnDestroy) return;
          py::gil_scoped_acquire gil;
          auto handler = PyBorrow(_handler);
          if (handler.is_none()) return;
          checkUserException([=]() { handler.attr("on_destroy")(); });
        }
        py::weakref _handler;
        bool _hasOnVisible, _hasOnDocked, _hasOnDestroy;
      };

      VoidFuture TaskPaneFrame_attach(
        ICustomTaskPane& self, const py::object& eventHandler, size_t hwnd)
      {
        return runExcelThread([
          &self, 
          hwnd, 
          handler = make_shared<PyTaskPaneHandler>(eventHandler)
        ]()
        {
          self.listen(handler);
          self.attach(hwnd);
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
        VoidFuture::bind(mod);

        py::class_<PyRibbonControl>(mod,
          "RibbonControl", R"(
            This object is passed to ribbon callback handlers to indicate which control  
            raised the callback.
          )")
          .def_readonly("id", 
            &PyRibbonControl::Id,
            "A string that represents the Id attribute for the control or custom menu item")
          .def_readonly("tag", 
            &PyRibbonControl::Tag,
            "A string that represents the Tag attribute for the control or custom menu item.");

        py::class_<ICustomTaskPane, shared_ptr<ICustomTaskPane>>(mod, 
          "TaskPaneFrame", R"(
            Manages Excel's underlying custom task pane object into which a python GUI can be
            drawn. It is unlikely that this object will need to be manipulated directly. Rather 
            use `xloil.gui.CustomTaskPane` which holds the python-side frame contents.

            The methods of this object are safe to call from any thread. COM must be used on 
            Excel's main thread, so the methods all wrap their calls to ensure to this happens.
          )")
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
          .def_property("position",
            TaskPane_getPosition,
            TaskPane_setPosition,
            R"(
              Gets/sets the dock position, one of: bottom, floating, left, right, top
            )")
          .def("com_control", 
            [](ICustomTaskPane& self, const char* binder)
            {
              py::gil_scoped_release noGil;
              return comToPy(*self.content(), binder);
            },
            R"(
              Gets the base COM control of the task pane. The ``lib`` used to provide
              COM support can be 'comtypes' or 'win32com' (default is win32com). This 
              method is only useful if a custom `progid` was specified during the task
              pane creation.
            )",
            py::arg("lib") = "")
          .def("attach", 
            &TaskPaneFrame_attach,
            R"( 
              Associates a `xloil.gui.CustomTaskPane` with this frame. Returns a future
              with no result.
            )",
            py::arg("handler"), 
            py::arg("hwnd"));

        py::class_<ComAddin>(mod, 
          "ExcelGUI", R"(
            An `ExcelGUI` wraps a COM addin which allows Ribbon customisation and creation
            of custom task panes. The methods of this object are safe to call from any thread;  
            however, since COM calls must be made on Excel's main thread, the methods schedule 
            those calls and return an *awaitable* future to the result. This could lead to deadlocks
            if the future's result is requested synchronously and, for example, one of Excel's event
            handlers is triggered. The object's properties do not return futures and are thread-safe.
          )")
          .def(py::init<py::object, py::object, py::object, bool>(),
            R"(
              Creates an `ExcelGUI` using the specified ribbon customisation XML
              and optionally connects it to Excel, ready for use.

              When the *ExcelGUI* object is deleted, it unloads the associated COM 
              add-in and so all Ribbon customisation and attached task panes.

              Parameters
              ----------

              ribbon: str
                  A Ribbon XML string, most easily created with a specialised editor.
                  The XML format is documented on Microsoft's website

              funcmap: Func[str -> callable] or Dict[str, callabe]
                  The ``funcmap`` mapper links callbacks named in the Ribbon XML to
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

              connect: bool
                  Defaults to True, meaning the object creation is blocking. If False
                  is passed, the object will not been fully constructed until the async
                  `connect` method is called.  In this case, no *ExcelGUI* methods can 
                  be called until the `connect` method has returned a result.

            )",
            py::arg("name") = py::none(),
            py::arg("ribbon") = py::none(),
            py::arg("funcmap") = py::none(),
            py::arg("connect" ) = true)
          .def("connect",
            &ComAddin::connect,
            R"(
              Connects the underlying COM addin to Excel, No other methods may be called 
              on a `ExcelGUI` object until it has been connected.

              This method is safe to call on an already-connected addin.
            )")
          .def("disconnect",
            &ComAddin::disconnect,
            R"(
              Unloads the underlying COM add-in and any ribbon customisation.  Avoid using
              connect/disconnect to modify the Ribbon as it is not perfomant. Rather hide/show
              controls with `invalidate` and the vibility callback.
            )")
          .def("invalidate",
            &ComAddin::invalidate,
            R"(
              Invalidates the specified control: this clears the cache of responses
              to callbacks associated with the control. For example, this can be
              used to hide a control by forcing its getVisible callback to be invoked,
              rather than using the cached value.

              If no control ID is specified, all controls are invalidated.
            )",
            py::arg("id") = "")
          .def("activate",
            &ComAddin::activate,
            R"(
              Activatives the ribbon tab with the specified id.  Returns False if
              there is no Ribbon or the Ribbon is collapsed.
            )",
            py::arg("id"))
          .def("_create_task_pane_frame",
            &ComAddin::createTaskPaneFrame,
            R"(
              Used internally to create a custom task pane window which can be populated
              with a python GUI.  Most users should use `attach_pane(...)` instead.

              A COM `progid` can be specified, but this will prevent displaying a python GUI
              in the task pane using the xlOil methods. This is a specialised use case.
            )",
            py::arg("name"), 
            py::arg("progid") = py::none(), 
            py::arg("window") = py::none())
          .def("attach_pane_async", 
            attachTaskPaneAsync,
            R"(
              Behaves as per `attach_pane`, but returns an *asyncio* coroutine. The
              `pane` argument may be an awaitable to a `CustomTaskPane`.
            )", 
            py::arg("pane"),
            py::arg("name") = py::none(),
            py::arg("window")=py::none(),
            py::arg("size")=py::none(),
            py::arg("visible")=true)
          .def("attach_pane",
            attachTaskPane,
            R"(
              Given task pane contents (which can be specified in several forms) this function
              creates a new task pane displaying those contents.

              Returns the instance of `CustomTaskPane`.  If one was passed as the 'pane' argument, 
              that is returned, if a *QWidget* was passed, a `QtThreadTaskPane` is created.

              Parameters
              ----------

              pane: CustomTaskPane (or QWidget type)
                  Can be an instance of `CustomTaskPane`, a type deriving from `QWidget` or
                  an instance of a `QWidget`. If a QWidget instance is passed, it must have 
                  been created on the Qt thread.

              name: 
                  The task pane name. Will be displayed above the task pane. If not provided,
                  the 'name' attribute of the task pane is used.

              window: 
                  A window title or `xloil.ExcelWindow` object to which the task pane should be
                  attached.  If None, the active window is used.

              size:
                  If provided, a tuple (width, height) used to set the initial pane size

              visible:
                  Determines the initial pane visibility. Defaults to True.
            )")
          .def("create_task_pane",
            createTaskPane,
            R"(
              Deprecated: use `attach_pane`. Note that `create_task_pane` tries to `find_task_pane`
              before creation whereas `attach_pane` does not.
            )", 
            py::arg("name"),
            py::arg("creator"),
            py::arg("window") = py::none(),
            py::arg("size") = py::none(),
            py::arg("visible") = true)
          .def_property_readonly("name",
            &ComAddin::name,
            "The name displayed in Excel's COM Addins window")
          .def_property_readonly("connected",
            &ComAddin::connected,
            "True if the a connection to Excel has been made");
      });
    }
  }
}