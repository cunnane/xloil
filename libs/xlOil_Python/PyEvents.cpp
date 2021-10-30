#include "PyEvents.h"
#include <xlOil/Events.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelRange.h>
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
        ArithmeticRef<T> operator()(T& x) const {
          return ArithmeticRef<T> { x };
        }
      };
    }

    template<class TEvent, bool, class F> class PyEvent {};

    // Specialisation to allow capture of the arguments to the event handler
    template<class TEvent, bool TAllowUserException, class R, class... Args>
    class PyEvent<TEvent, TAllowUserException, std::function<R(Args...)>>
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

      PyEvent& add(const py::object& obj)
      {
        if (_handlers.empty())
          _coreEventHandler = _event += [this](Args... args) { this->fire(args...); };

        // We use a weakref to avoid dangling pointers to event handlers
        // the callback calls this->remove(ptr)
        _handlers.push_back(py::weakref(obj, _refRemover));
        return *this;
      }

      PyEvent& remove(const py::object& obj)
      {
        _handlers.remove(obj);
        // Unhook ourselves from the core for efficiency if there are no handlers
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
            auto* handler = PyWeakref_GET_OBJECT(h.ptr());
            // See above for the purpose of ReplaceArithmeticRef
            if (handler != Py_None)
              py::cast<py::function>(handler)(ReplaceArithmeticRef<Args>()(args)...);
          }
        }
        catch (const py::error_already_set& e)
        {
          // Avoid recursion if we actually are Event_PyUserException!
          if constexpr(TAllowUserException)
            Event_PyUserException().fire(e.type(), e.value(), e.trace());
          XLO_ERROR("During Event {0}: {1}", _event.name(), e.what());
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("During Event {0}: {1}", _event.name(), e.what());
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

      template<class T>
      void bindEvent(py::module& mod, T* event, const char* name)
      {
        const auto& instances = py::detail::get_internals().registered_types_cpp;
        const auto found = instances.find(std::type_index(typeid(T)));
        if (found == instances.end())
        {
          py::class_<T>(mod, (string(name) + "_Type").c_str())
            .def("__iadd__", &T::add)
            .def("__isub__", &T::remove)
            .def("handlers", &T::handlers)
            .def("clear", &T::clear);
        }
        mod.add_object(name, py::cast(event, py::return_value_policy::take_ownership));
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
        Event::allowEvents(value);
      }
      static int theBinder = addBinder([](pybind11::module& mod)
      {
        auto eventMod = mod.def_submodule("event");
        eventMod.def("allow", []() {setAllowEvents(true); });
        eventMod.def("pause", []() {setAllowEvents(false); });

        bindArithmeticRef<bool>(eventMod);

#define XLO_PY_EVENT(r, _, NAME) \
        bindEvent(eventMod, makeEvent(xloil::Event::NAME()), BOOST_PP_STRINGIZE(NAME));

        BOOST_PP_SEQ_FOR_EACH(XLO_PY_EVENT, _, XLOIL_STATIC_EVENTS)
#undef XLO_PY_EVENT

        bindEvent(eventMod, makeEventNoUserExcept(Event_PyUserException()), "UserException");
        bindEvent(eventMod, makeEventNoUserExcept(Event_PyBye()), "PyBye");
      });
    }
  }
}