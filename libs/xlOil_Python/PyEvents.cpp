#include "PyEvents.h"
#include <xlOil/Events.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelRange.h>
#include "PyHelpers.h"
#include "InjectedModule.h"
#include <list>

namespace py = pybind11;
using std::unordered_map;
using std::shared_ptr;
using std::string;

namespace xloil 
{
  namespace Python 
  {
    namespace
    {
      template <class T>
      struct ArithmeticRef
      {
        T& value;
      };

      template<class T, bool = std::is_arithmetic<std::remove_reference_t<T>>::value>
      struct ReplaceAritmeticRef
      {
        T operator()(T x) const { return x; }
      };

      template<class T>
      struct ReplaceAritmeticRef<T&, true>
      {
        auto operator()(T& x) const {
          return ArithmeticRef<T> { x };
        }
      };

      template<class T>
      auto magic(T x)
      {
        return ReplaceAritmeticRef<T>()(x);
      }
    }

    template<class TEvent, class F> class PyEvent2 {};

    template<class TEvent, class R, class... Args>
    class PyEvent2<TEvent, std::function<R(Args...)>>
    {
    public:
      PyEvent2(TEvent& event) 
        : _event(event) 
      {
        _refRemover = py::cpp_function([this](py::weakref& ref) { this->remove(ref); });
      }

      ~PyEvent2()
      {
        if (!_handlers.empty())
          _event -= _coreEventHandler;
      }
      PyEvent2& add(const py::object& obj)
      {
        if (_handlers.empty())
          _coreEventHandler = _event += [this](Args... args) { this->fire(args...); };
        _handlers.push_back(py::weakref(obj, _refRemover));
        return *this;
      }
      PyEvent2& remove(const py::object& obj)
      {
        _handlers.remove(obj); // is py::object equality going to work?
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
            if (handler != Py_None)
              py::cast<py::function>(handler)(magic<Args>(args)...);
          }
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("During Event: {0}", e.what());
        }
      }

    /*  void unloadModule(const py::module& mod)
      {
        for (auto i = _handlers.begin(); i != _handlers.end();)
        {
          if (mod == i->attr("__module__"))
            _handlers.erase(i++);
          else ++i;
        }
        if (_handlers.empty())
          _event -= _coreEventHandler;
      }*/

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
        return std::make_shared<PyEvent2<TEvent, typename TEvent::handler>>(event);
      }

      template<class T>
      void bindEvent(py::module& mod, const shared_ptr<T>& event, const char* name)
      {
        const auto& instances = py::detail::get_internals().registered_types_cpp;
        const auto found = instances.find(std::type_index(typeid(T)));
        if (found == instances.end())
        {
          py::class_<T, shared_ptr<T>>(mod, (string(name) + "_Type").c_str())
            .def("__iadd__", &T::add)
            .def("__isub__", &T::remove)
            .def("handlers", &T::handlers);
        }
        mod.add_object(name, py::cast(event));
      }

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

      static int theBinder = addBinder([](pybind11::module& mod)
      {
        auto eventMod = mod.def_submodule("event");
        bindArithmeticRef<bool>(eventMod);

#define XLO_PY_EVENT(NAME) \
        auto NAME = makeEvent(xloil::Event::NAME()); \
        bindEvent(eventMod, NAME, #NAME);
        XLO_PY_EVENT(AfterCalculate);
        XLO_PY_EVENT(CalcCancelled);
        XLO_PY_EVENT(NewWorkbook);
        XLO_PY_EVENT(SheetSelectionChange);
        XLO_PY_EVENT(SheetBeforeDoubleClick);
        XLO_PY_EVENT(SheetBeforeRightClick);
        XLO_PY_EVENT(SheetActivate);
        XLO_PY_EVENT(SheetDeactivate);
        XLO_PY_EVENT(SheetCalculate);
        XLO_PY_EVENT(SheetChange);
        XLO_PY_EVENT(WorkbookOpen);
        XLO_PY_EVENT(WorkbookActivate);
        XLO_PY_EVENT(WorkbookDeactivate);
        XLO_PY_EVENT(WorkbookAfterClose);
        XLO_PY_EVENT(WorkbookBeforeSave);
        XLO_PY_EVENT(WorkbookBeforePrint);
        XLO_PY_EVENT(WorkbookNewSheet);
        XLO_PY_EVENT(WorkbookAddinInstall);
        XLO_PY_EVENT(WorkbookAddinUninstall);
#undef XLO_PY_EVENT
      });
    }
  }
}