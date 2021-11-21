#include "PyHelpers.h"
#include "PyEvents.h"
#include <CTPL/ctpl_stl.h>
#include <xlOil/Log.h>

namespace xloil
{
  namespace Python
  {
    class EventLoop
    {
      ctpl::thread_pool _thread;
      pybind11::object _eventLoop;
      pybind11::object _callSoonFunction;
      std::shared_ptr<const void> _shutdownHandler;

    public:
      EventLoop()
        : EventLoop([]() {})
      {}

      EventLoop(std::function<void()> init)
        : _thread(1)
      {
#ifdef _DEBUG
        if (PyGILState_Check() == 1)
          XLO_THROW("Release GIL before constructing an EventLoop");
#endif

        _thread.push([this, init](int) mutable
        {
          try
          {
            init();
            pybind11::gil_scoped_acquire getGil;
            getGil.inc_ref();
            _eventLoop = pybind11::module::import("asyncio").attr("new_event_loop")();
            _callSoonFunction = pybind11::module::import("xloil.register").attr("_loop_call_threadsafe");
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Failed to initialise python worker thread: {0}", e.what());
          }
        }).get();

        _thread.push([this](int) mutable
        {
          try
          {
            pybind11::gil_scoped_acquire getGil;
            getGil.dec_ref();
            _eventLoop.attr("run_forever")();
            _eventLoop = pybind11::object();
            _callSoonFunction = pybind11::object();
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Failed to initialise python worker thread: {0}", e.what());
          }
        });
      }
      ~EventLoop()
      {
        if (active())
          shutdown();
      }
      template <class...Args>
      void callback(const pybind11::object& func, Args&&... args)
      {
        if (!active())
          return;
        _callSoonFunction(_eventLoop, func, std::forward<Args>(args)...);
      }
      template <class...Args>
      void callback(const char* module, const char* func, Args&&... args)
      {
        callback(py::module::import(module).attr(func), std::forward<Args>(args)...);
      }
      bool active()
      {
        return _thread.size() > 0;
      }
      void stop()
      {
        pybind11::gil_scoped_acquire getGil;
        callback(_eventLoop.attr("stop"));
      }
      void shutdown()
      {
        _shutdownHandler.reset();
        stop();
        _thread.stop();
      }
      auto loop() { return _eventLoop; }
    };
  }
}