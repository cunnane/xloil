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
        //pybind11::gil_scoped_release releaseGil;

        _thread.push([this, init](int) mutable
        {
          try
          {
            init();
            pybind11::gil_scoped_acquire getGil;
            getGil.inc_ref();
            _eventLoop = pybind11::module::import("asyncio").attr("new_event_loop")();
            _callSoonFunction = _eventLoop.attr("call_soon_threadsafe");
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

        _shutdownHandler = std::static_pointer_cast<const void>(
          Event_PyBye().bind([self = this]
          {
            self->shutdown();
          }));
      }
      template <class...Args>
      void callback(const pybind11::object& func, Args&&... args)
      {
        if (!active())
          return;
        _callSoonFunction(func, std::forward<Args>(args)...);
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