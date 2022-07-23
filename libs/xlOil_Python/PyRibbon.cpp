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

      using VoidFuture = PyFuture<void, void>;
      using CTPFuture = PyFuture<shared_ptr<ICustomTaskPane>>;

      class ComAddin
      {
      public:
        ComAddin(wstring&& name)
          : _name(name)
        {}

        ~ComAddin()
        {
          py::gil_scoped_release noGil;
          runExcelThread([this]() { _addin->disconnect(); }).get();
          _addin.reset();
        }

        VoidFuture connect(const py::object& xml, const py::object& funcmap)
        {
          return runExcelThread([
              this,
              xmlStr = pyToWStr(xml),
              mapper = funcmap.is_none() ? IComAddin::RibbonMap() : makeRibbonNameMapper(funcmap)]
              () mutable
            {
              if (!_addin)
                _addin = makeComAddin(_name.c_str(), nullptr);
              _addin->connect(xmlStr.c_str(), mapper);
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
          auto progIdStr = progId.is_none() ? wstring() : pyToWStr(progId).c_str();
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
        std::atomic<bool> _connected;
      };
    
      auto attachTaskPane(
        const py::object& comAddin,
        const py::object& name,
        const py::object& pane,
        const py::object& window,
        const py::object& size,
        const py::object& visible)
      {
        auto attach = py::module::import("xloil.gui").attr("_attach_task_pane");
        return attach(comAddin, name, pane, window, size, visible);
      }

      class PyTaskPaneHandler : public ICustomTaskPaneEvents
      {
      public:
        PyTaskPaneHandler(const py::object& eventHandler)
          : _handler(eventHandler)
        {
          //TODO: check upfront which functions are implemented and do not call if not!
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

      void addPaneEventHandler(ICustomTaskPane& self, const py::object& eventHandler, size_t hwnd)
      {
        runExcelThread([&self, hwnd, handler = make_shared<PyTaskPaneHandler>(eventHandler)]()
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
            Excel's underlying custom task pane object into which a python GUI can be drawn.
            It is unlikely that this object will need to be manipulated directly. Rather use
            `xloil.gui.CustomTaskPane` which holds the python-side frame contents.

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
            &addPaneEventHandler, 
            R"( 
              Associates a `xloil.gui.CustomTaskPane` with this frame
            )",
            py::arg("handler"), 
            py::arg("hwnd"));

        py::class_<ComAddin>(mod, 
          "ExcelGUI", R"(
            Creating an ExcelGUI creates a COM addin which allows Ribbon customisation and creation
            of custom task panes. The methods of this object are safe to call from any thread;  
            however, since COM calls must be made on Excel's main thread, the methods schedule 
            those calls and return an *awaitable* future to the result. This could lead to deadlocks
            if the future's result is requested synchronously if, for example, one of Excel's event
            handlers is triggered.  The object's properties do not return futures and are thread-safe.

            *ExcelGUI* methods cannot be called until the future created by its *connect* method 
            has returned a result.
          )")
          .def(py::init<wstring&&>(), py::arg("name"))
          .def("connect",
            &ComAddin::connect,
            R"(
              Connects this COM addin to Excel, optionally specifying Ribbon XML and callbacks.
              The Ribbon specification can only be modified on connection. No other methods may
              be called on a *ExcelGUI* object until it has been connected.

              This method is safe to call on an already connected addin.
            )",
            py::arg("xml")="", 
            py::arg("func_names")=py::none())
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
              with a python GUI.  Most users should use `create_task_pane(...)` instead.

              A COM `progid` can be specified, but this will prevent displaying a python GUI
              in the task pane using the xlOil methods. This is a specialised use case.
            )",
            py::arg("name"), 
            py::arg("progid") = py::none(), 
            py::arg("window") = py::none())
          .def("attach_pane", 
            attachTaskPane,
            R"(
              
              Parameters
              ----------

              name: 
                  The task pane name. Will be displayed above the task pane.

              pane: CustomTaskPane (or QWidget type)
                  If a QWidget instance is passed, it must have been created on the Qt thread
                  or core dumps will ensue.

              window: 
                  A window title or `xloil.ExcelWindow` object to which the task pane should be
                  attached.  If None, the active window is used.

              size:
                  If provided, a tuple (width, height) used to set the initial pane size

              visible:
                  Determines the initial pane visibility. Defaults to True.
            )", 
            py::arg("name"), 
            py::arg("pane"),
            py::arg("window")=py::none(),
            py::arg("size")=py::none(),
            py::arg("visible")=true)
          .def_property_readonly("name", 
            &ComAddin::name)
          .def_property_readonly("connected",
            &ComAddin::connected);
      });
    }
  }
}