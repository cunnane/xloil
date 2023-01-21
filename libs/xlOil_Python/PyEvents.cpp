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

namespace py = pybind11;
using std::unordered_map;
using std::shared_ptr;
using std::string;

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
    
    struct IPyEvent
    {
      virtual ~IPyEvent() {}
      virtual IPyEvent& add(const py::object& obj) = 0;
      virtual IPyEvent& remove(const py::object& obj) = 0;
      virtual py::tuple handlers() const = 0;
      virtual void clear() = 0;
    };

    template<class TEvent, bool, class F> class PyEvent {};

    // Specialisation to allow capture of the arguments to the event handler
    template<class TEvent, bool TAllowUserException, class R, class... Args>
    class PyEvent<TEvent, TAllowUserException, std::function<R(Args...)>> : public IPyEvent
    {
    public:
      PyEvent(TEvent& event) 
        : _event(event) 
      {
        // This is called by weakref when the ref count goes to zero
        _refRemover = py::cpp_function([this](py::weakref& ref) { this->remove(ref); });
      }

      ~PyEvent()
      {
        if (!_handlers.empty())
          _event -= _coreEventHandler;
      }

      IPyEvent& add(const py::object& obj)
      {
        if (_handlers.empty())
          _coreEventHandler = _event += [this](Args... args) { this->fire(args...); };

        XLO_INFO(L"Event {} added handler {}", _event.name(), (void*)obj.ptr());

        // We use a weakref to avoid dangling pointers to event handlers.
        // For bound methods, we need the WeakMethod class to avoid them
        // being immediately deleted.

        if (py::hasattr(obj, "__self__"))
          _handlers.push_back(py::module::import("weakref").attr("WeakMethod")(obj));
        else
          _handlers.push_back(py::weakref(obj, _refRemover));

        return *this;
      }

      IPyEvent& remove(const py::object& obj)
      {
        _handlers.remove(obj);
        // Unhook ourselves from the core for efficiency if there are no handlers
        XLO_INFO(L"Event {} removed handler {}", _event.name(), (void*)obj.ptr());
        if (_handlers.empty())
          _event -= _coreEventHandler;
        return *this;
      }

      py::tuple handlers() const
      {
        py::tuple result(_handlers.size());
        auto i = 0;
        for (auto iHandler = _handlers.begin(); iHandler != _handlers.end(); ++iHandler)
          result[i] = *iHandler;
        return result;
      }

      void fire(Args... args) const
      {
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

      void clear()
      {
        _event.clear();
      }

    private:
      TEvent& _event;
      std::list<py::weakref> _handlers;
      typename TEvent::handler_id _coreEventHandler;
      py::function _refRemover;
    };

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
          py::cast(event, py::return_value_policy::take_ownership));
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
        const auto name = string(typeid(T).name()) + "Ref";
        using RefType = ArithmeticRef<T>;
        py::class_<RefType>(mod, name.c_str())
          .def_property("value",
            [](const RefType& self) { return self.value; },
            [](RefType& self, T val) { self.value = val; });
      }

      void setAllowEvents(bool value)
      {
        py::gil_scoped_release releaseGil;
        runExcelThread([=]() { excelApp().setEnableEvents(value); });
      }

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        auto eventMod = mod.def_submodule("event");
        eventMod.doc() = R"(
          A module containing event objects which can be hooked to receive events driven by 
          Excel's UI. The events correspond to COM/VBA events and are described in detail
          in the Excel Application API.
        
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
                :ref:`xloil.Range`
    
          Python-only Events
          ------------------

          These events specific to python and not documented in the Core documentation:

            * PyBye:
                Fired just before xlOil finalises its embedded python interpreter. 
                All python and xlOil functionality is still available. This event is useful 
                to stop threads as it is called before threading module teardown, whereas 
                python's `atexit` is called afterwards. Has no parameters.
            * UserException:
                Fired when an exception is raised in a user-supplied python callback, 
                for example a GUI callback or an RTD publisher. Has no parameters.

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
          []() { setAllowEvents(true); },
          R"(
            Resumes Excel's event handling after a pause.  Equivalent to VBA's
            `Application.EnableEvents = True` or `xlo.app().enable_events = True` 
          )" );

        eventMod.def("pause", 
          []() { setAllowEvents(false); },
          R"(
            Pauses Excel's event handling. Equivalent to VBA's 
            `Application.EnableEvents = False` or `xlo.app().enable_events = False` 
          )");

        bindArithmeticRef<bool>(eventMod);

        py::class_<IPyEvent>(eventMod, "Event")
          .def("__iadd__", &IPyEvent::add)
          .def("__isub__", &IPyEvent::remove)
          .def_property_readonly("handlers", &IPyEvent::handlers)
          .def("clear", &IPyEvent::clear);

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
      });
    }

    void raiseUserException(const pybind11::error_already_set& e)
    {
      Event_PyUserException().fire(e.type(), e.value(), e.trace());
    }
  }
}