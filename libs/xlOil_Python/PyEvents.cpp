#include "PyEvents.h"
#include <xlOil/Events.h>
#include <xlOil/Log.h>
#include <xlOil/Range.h>
#include <xlOil/AppObjects.h>
#include "PyHelpers.h"
#include "PyCore.h"
#include <list>
#include <boost/preprocessor/seq/for_each.hpp>
#include <boost/preprocessor/stringize.hpp>
#include <xloil/Interface.h>
#include <filesystem>

namespace py = pybind11;
using std::unordered_map;
using std::shared_ptr;
using std::string;
using std::wstring;
namespace fs = std::filesystem;

namespace std 
{
  bool operator==(const py::weakref& lhs, const py::weakref& rhs)
  {
    return lhs.is(rhs);
  }
}

namespace xloil 
{
  namespace Python 
  {
    XLOIL_DEFINE_EVENT(Event_PyBye);
    XLOIL_DEFINE_EVENT(Event_PyUserException);

    class IPyEvent;

    // GLOBALS
    namespace {
      std::unordered_map<wstring, shared_ptr<IPyEvent>> theDirChangeEvents;
      auto theDirChangeEventsCleanup = Event_PyBye().bind([]() { theDirChangeEvents.clear(); });

      std::atomic<bool> theEventsAreEnabled = true;
    }

    namespace
    {
      /// <summary>
      /// This struct is designed to hold a reference to an arithmetic type so
      /// it can be modified in Python, otherwise arithmetic types are immutable.
      /// </summary>
      template <class T>
      struct ArithmeticRef
      {
        T& value;
      };

      /// <summary>
      /// Some template magic which I don't fully understand to replace a 
      /// non-const ref to an arithmetic type with ArithmeticRef.
      /// </summary>
      template<class T, bool = std::is_arithmetic<std::remove_reference_t<T>>::value>
      struct ReplaceArithmeticRef
      {
        auto operator()(T x) const { return x; }
      };
      template<class T>
      struct ReplaceArithmeticRef<const T&, false>
      {
        const T& operator()(const T& x) const { return x; }
      };
      template<class T>
      struct ReplaceArithmeticRef<T&, true>
      {
        auto operator()(T& x) const {
          return ArithmeticRef<T> { x };
        }
      };
    }
    
    class IPyEvent
    {
    public:
      virtual ~IPyEvent()
      {
        if (!_handlers.empty())
        {
          py::gil_scoped_acquire getGil;
          unbind();
          _handlers.clear();
        }
      }

      virtual const wchar_t* name() const = 0;

      py::tuple handlers() const
      {
        py::tuple result(_handlers.size());
        auto i = 0;
        for (auto iHandler = _handlers.begin(); iHandler != _handlers.end(); ++iHandler)
          result[i] = *iHandler;
        return result;
      }

      IPyEvent& add(const py::object& obj)
      {
        // This is called by weakref when the ref count goes to zero
        py::object callback;
        if (_handlers.empty())
        {
          bind();
          callback = py::cpp_function([this](py::object& ref) { this->remove(ref); });
        }
        else
          callback = _handlers.back().attr("__callback__");

        XLO_INFO(L"Event {} added handler {}", name(), (void*)obj.ptr());

        // We use a weakref to avoid dangling pointers to event handlers.
        // For bound methods, we need the WeakMethod class to avoid them
        // being immediately deleted.
        if (py::hasattr(obj, "__self__"))
          _handlers.push_back(py::module::import("weakref").attr("WeakMethod")(obj, callback));
        else
          _handlers.push_back(py::weakref(obj, callback));

        return *this;
      }

      IPyEvent& remove(const py::object& obj)
      {
        _handlers.remove(obj);
        // Unhook ourselves from the core for efficiency if there are no handlers
        XLO_INFO(L"Event {} removed handler {}", name(), (void*)obj.ptr());
        if (_handlers.empty())
        {
          unbind();
          XLO_DEBUG(L"No more python handlers for event {}", name());
        }
        return *this;
      }

      void clear()
      {
        unbind();
        _handlers.clear();
      }

      virtual void bind() = 0;
      virtual void unbind() {} // Why?

    protected:
      IPyEvent()
      {}

      std::list<py::weakref> _handlers;
    };

    template<class TEvent, bool, class F> class PyEvent {};

    // Specialisation to allow capture of the arguments to the event handler
    template<class TEvent, bool TAllowUserException, class R, class... Args>
    class PyEvent<TEvent, TAllowUserException, std::function<R(Args...)>> : public IPyEvent
    {
    public:
      PyEvent(TEvent& event)
        : _event(event)
        , _coreEventHandler(nullptr)
      {}

      void unbind() override
      {
        _event -= _coreEventHandler;
      }

      const wchar_t* name() const override
      {
        return  _event.name().c_str();
      }

      void bind() override
      {
        _coreEventHandler = _event += [this](Args... args) { this->fire(args...); };
      }

      void fire(Args... args) const
      {
        if (!theEventsAreEnabled)
          return;

        try
        {
          py::gil_scoped_acquire get_gil;
          for (auto& h : _handlers)
          {
            auto handler = h();
            // See above for the purpose of ReplaceArithmeticRef
            if (!handler.is_none())
              handler(ReplaceArithmeticRef<Args>()(args)...);
          }
        }
        catch (const py::error_already_set& e)
        {
          // Avoid recursion if we actually are Event_PyUserException!
          if constexpr(TAllowUserException)
            Event_PyUserException().fire(e.type(), e.value(), e.trace());
          XLO_ERROR(L"During Event {0}: {1}", _event.name(), utf8ToUtf16(e.what()));
        }
        catch (const std::exception& e)
        {
          XLO_ERROR(L"During Event {0}: {1}", _event.name(), utf8ToUtf16(e.what()));
        }
      }

    private:
      TEvent& _event;
      typename TEvent::handler_id _coreEventHandler;
    };

    class PyDirectoryWatch: public IPyEvent
    {
    public:
      PyDirectoryWatch(
          const std::wstring_view& objectPath, 
          Event::FileAction action, 
          bool subDirs)
        : _coreEventHandler(nullptr)
      {
        auto path = fs::path(objectPath);
        if (fs::is_directory(path))
        {
          _directory = path;
          _watchSubDirs = subDirs;
        }
        else
        {
          _filename = path.filename();
          _directory = path.remove_filename();
          _watchSubDirs = false;
        }
        _name = _directory / _filename;
        _action = action;
      }

      void unbind() override
      {
        _coreEventHandler.reset();
      }

      const wchar_t* name() const override
      {
        return _name.c_str(); // Include the filename silly
      }

      void bind() override
      {
        _coreEventHandler = Event::DirectoryChange(_directory, _watchSubDirs)->bind(
          [this](const wchar_t* directory, const wchar_t* filename, Event::FileAction action) 
          { 
            this->fire(directory, filename, action);
          });
      }

      void fire(const wchar_t* directory, const wchar_t* filename, Event::FileAction action) const
      {
        if (!theEventsAreEnabled)
          return;

        // Check out action and filename (if specified) match, if not, we don't 
        // need to acquire the gil.
        if (action != _action || (!_filename.empty() && _filename != filename))
          return;

        const auto target = fs::path(directory) / filename;

        try
        {
          // Note if _handlers is empty the event will already have been unbound
          py::gil_scoped_acquire get_gil;
          for (auto& h : _handlers)
          {
            auto handler = h(); // Get strong ref
            if (!handler.is_none())
              handler(target.c_str());
          }
        }
        catch (const py::error_already_set& e)
        {
          Event_PyUserException().fire(e.type(), e.value(), e.trace());
          XLO_ERROR(L"During Event {0}: {1}", name(), utf8ToUtf16(e.what()));
        }
        catch (const std::exception& e)
        {
          XLO_ERROR(L"During Event {0}: {1}", name(), utf8ToUtf16(e.what()));
        }
      }

    private:
      fs::path _directory;
      wstring _filename;
      wstring _name;
      Event::FileAction _action;
      bool _watchSubDirs;
      shared_ptr<const void> _coreEventHandler;
    };

    void setAllowPyEvents(bool value)
    {
      theEventsAreEnabled = value;
    }

    namespace
    {
      template<class TEvent>
      auto makeEvent(TEvent& event)
      {
        return new PyEvent<TEvent, true, typename TEvent::handler>(event);
      }

      /// <summary>
      /// Binds an event which does not fire the UserException handler - useful
      /// to avoid circular event calls
      /// </summary>
      template<class TEvent>
      auto makeEventNoUserExcept(TEvent& event)
      {
        return new PyEvent<TEvent, false, typename TEvent::handler>(event);
      }

      void bindEvent(
        py::module& mod, 
        const wchar_t* name, 
        IPyEvent* event)
      {
        auto u8name = utf16ToUtf8(name);
        mod.add_object(u8name.c_str(), 
          py::cast(shared_ptr<IPyEvent>(event), py::return_value_policy::take_ownership));
      }

      /// <summary>
      /// ArithmeticRef is designed to hold a reference to an arithmetic type so
      /// it can be modified in Python, otherwise arithmetic types are immutable.
      /// Python doesn't allow override of the '=' operator so we have to just
      /// expose the 'value property'
      /// </summary>
      template<class T>
      void bindArithmeticRef(py::module& mod)
      {
        const auto name = "_" + string(typeid(T).name()) + "_ref";
        using RefType = ArithmeticRef<T>;
        py::class_<RefType>(mod, name.c_str())
          .def_property("value",
            [](const RefType& self) { return self.value; },
            [](RefType& self, T val) { self.value = val; });
      }      
      
      void setEnableEvents(bool value, bool excelEvents)
      {
        py::gil_scoped_release releaseGil;
        XLO_DEBUG("Events enabled: Excel={}, xlOil={}", !excelEvents || value, value);
        setAllowPyEvents(value);
        if (excelEvents)
          runExcelThread([=]() { 
            thisApp().setEnableEvents(value); 
          });
      }

      auto getDirectoryChangeEvent(const wstring& path, wstring& action, bool subDirs)
      {
        using Event::FileAction;
        toLower(action);
        FileAction act;
        if (action == L"add")
          act = FileAction::Add;
        else if (action == L"delete")
          act = FileAction::Delete;
        else if (action == L"modify")
          act = FileAction::Modify;
        else
          throw py::value_error("action");

        auto key = path + L'_' + action;
        // TODO: actually use the "action" parameter!
        auto found = theDirChangeEvents.find(key);
        if (found != theDirChangeEvents.end())
          return found->second;
        
        auto result = shared_ptr<IPyEvent>(new PyDirectoryWatch(path, act, subDirs));
        theDirChangeEvents[key] = result;
        return result;
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        auto eventMod = mod.def_submodule("event");
        eventMod.doc() = R"(
          A module containing event objects which can be hooked to receive events driven by 
          Excel's UI. The events correspond to COM/VBA events and are described in detail
          in the Excel Application API. The naming convention (including case) of the VBA events
          has been preserved for ease of search.
          
        
          See :ref:`Events:Introduction` and 
          `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_

          Using the Event Class
          ---------------------

              * Events are hooked using `+=`, e.g.

              ::
              
                  event.NewWorkbook += lambda wb: print(wb_name)

              * Events are unhooked using `-=` and passing a reference to the handler function

              ::

                  event.NewWorkbook += foo
                  event.NewWorkbook -= foo

              * You should not return anything from an event handler

              * Each event has a `handlers` property listing all currently hooked handlers

              * Where an event has reference parameter, for example the `cancel` bool in
                `WorkbookBeforePrint`, you need to set the value using `cancel.value=True`.
                This is because python does not support reference parameters for primitive types.

                ::

                    def no_printing(wbName, cancel):
                      cancel.value = True
                    xlo.event.WorkbookBeforePrint += no_printing

              * Workbook and worksheet names are passed a string, Ranges as passed as a 
                :any:`xloil.Range`
    
          Python-only Events
          ------------------

          These events are specific to python and not noted in the Core documentation:

            * PyBye:
                Fired just before xlOil finalises its embedded python interpreter. 
                All python and xlOil functionality is still available. This event is useful 
                to stop threads as it is called before threading module teardown, whereas 
                python's `atexit` is called afterwards. Has no parameters.
            * UserException:
                Fired when an exception is raised in a user-supplied python callback, 
                for example a GUI callback or an RTD publisher. Has no parameters.
            * file_change:
                This is a special parameterised event, see the separate documentation
                for this function.

          Examples
          --------

          ::

              def greet(workbook, worksheet):
                  xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

              xlo.event.WorkbookNewSheet += greet
              ...
              xlo.event.WorkbookNewSheet -= greet
              
              print(xlo.event.WorkbookNewSheet.handlers) # Should be empty


          ::

              def click_handler(sheet_name, target, cancel):
                  xlo.worksheets[sheet_name]['A5'].value = target.address()
    
              xlo.event.SheetBeforeDoubleClick += click_handler

        )";

        eventMod.def("allow", 
          [](bool excel) { setEnableEvents(true, excel); },
          R"(
            Resumes event handling after a previous call to *pause*.

            If *excel* is True (the default), also calls `Application.EnableEvents = True`
            (equivalent to `xlo.app().enable_events = True`)
          )",
          py::arg("excel") = true);

        eventMod.def("pause", 
          [](bool excel) { setEnableEvents(false, excel); },
          R"(
            Stops all xlOil event handling - any executing handlers will complete but
            no further handlers will fire.

            If *excel* is True (the default), also calls `Application.EnableEvents = False`
            (equivalent to `xlo.app().enable_events = False`)
          )",
          py::arg("excel")=true);

        bindArithmeticRef<bool>(eventMod);

        py::class_<IPyEvent, shared_ptr<IPyEvent>>(eventMod, "Event")
          .def("__iadd__", &IPyEvent::add)
          .def("__isub__", &IPyEvent::remove)
          .def(
            "add", 
            &IPyEvent::add,
            R"(
              Registers an event handler callback with this event, equivalent to
              `event += handler`
            )", py::arg("handler"))
          .def(
            "remove", 
            &IPyEvent::remove,
            R"(
              Deregisters an event handler callback with this event, equivalent to
              `event -= handler`
            )", py::arg("handler"))
          .def_property_readonly(
            "handlers", 
            &IPyEvent::handlers,
            "The tuple of handlers registered for this event. Read-only.")
          .def(
            "clear", 
            &IPyEvent::clear,
            "Removes all handlers from this event");

        // TODO: how to set doc string for each event?
#define XLO_PY_EVENT(r, _, NAME) \
        bindEvent(eventMod, XLO_WSTR(NAME), makeEvent(xloil::Event::NAME()));

        BOOST_PP_SEQ_FOR_EACH(XLO_PY_EVENT, _, XLOIL_STATIC_EVENTS)
#undef XLO_PY_EVENT

        bindEvent(eventMod, 
          L"UserException",
          makeEventNoUserExcept(Event_PyUserException()));

        bindEvent(eventMod, 
          L"PyBye",
          makeEventNoUserExcept(Event_PyBye()));

        eventMod.def(
          "file_change", 
          getDirectoryChangeEvent, 
          R"(
            This function returns an event specific to the given path and action; the 
            event fires when a watched file or directory changes.  The returned event 
            can be hooked in the usual way using `+=` or `add`. Calling this function 
            with same arguments always returns a reference to the same event object.

            The handler should take a single string argument: the name of the file or 
            directory which changed.

            The event runs on a background thread.

            Parameters
            ----------

            path: str
               Can be a file or a directory. If *path* points to a directory, any change
               to files in that directory, will trigger the event. Changes to the specified 
               directory itself will not trigger the event.

            action: str ["add", "remove", "modify"], default "modify"
               The event will only fire when this type of change is detected:

                 *modify*: any update which causes a file's last modified time to change
                 *remove*: file deletion
                 *add*: file creation

               A file rename triggers *remove* followed by *add*.

            subdirs: bool (true)
               including in subdirectories,
          )",
          py::arg("path"), 
          py::arg("action") = "modify",
          py::arg("subdirs") = true);
      });
    }

    void raiseUserException(const pybind11::error_already_set& e)
    {
      // Acquire gil here as if debug logging is enabled, the event base class
      // will try to write out the event parameters as strings.
      py::gil_scoped_acquire gilAcquired;
      Event_PyUserException().fire(e.type(), e.value(), e.trace());
    }
  }
}