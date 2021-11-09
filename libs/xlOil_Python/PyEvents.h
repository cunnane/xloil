#pragma once
#include <xlOil/Events.h>
namespace pybind11 {
  class object; class error_already_set;
}
namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// An event triggered when the Python plugin is about to close
    /// but before the Python interpreter is stopped.
    /// </summary>
    Event::Event<void(void)>& Event_PyBye();

    /// <summary>
    /// Fired when an exception occurs in user-defined code and passes the
    /// (type, value, traceback) argument triple as per sys.exec_info().
    /// Allows custom handling of user exceptions, e.g. opening a debugger.
    /// </summary>
    Event::Event<void(
      const pybind11::object&, 
      const pybind11::object&, 
      const pybind11::object&)>&
        Event_PyUserException();

    void raiseUserException(const pybind11::error_already_set& e);
    
    template<class F, class... Args>
    auto checkUserException(F&& f, Args... args)
    {
      try
      {
        return f(std::forward<Args>(args)...);
      }
      catch (const pybind11::error_already_set& e)
      {
        raiseUserException(e);
        throw;
      }
    }
  }
}