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
      pybind11::object _pumpFunction;
      std::atomic<bool> _stopped;

    public:
      unsigned asyncioTimeout;
      unsigned sleepTime;

      EventLoop(unsigned asyncioTimeout_ = 200, unsigned sleepTime_ = 200)
        : _thread(1) 
        , _stopped(false)
        , asyncioTimeout(asyncioTimeout_)
        , sleepTime(sleepTime_)
      {
#ifdef _DEBUG
        if (PyGILState_Check() == 1)
          XLO_THROW("Release GIL before constructing an EventLoop");
#endif

        _thread.push([this](int) mutable
        {
          try
          {
            pybind11::gil_scoped_acquire getGil;
            // Create a hanging ref to the python thread state, to avoid all context vars,
            // particularly the asyncio loop, being deleted at the end of this function.
            getGil.inc_ref();

            _eventLoop = pybind11::module::import("asyncio").attr("new_event_loop")();
            _pumpFunction = pybind11::module::import("xloil.importer").attr("_pump_message_loop");
            _callSoonFunction = _eventLoop.attr("call_soon_threadsafe");
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Failed to initialise python event loop: {0}", e.what());
          }
        }).get();

        if (!_eventLoop || _eventLoop.is_none() || !_callSoonFunction || _callSoonFunction.is_none())
          XLO_THROW("Failed starting event loop");

        _thread.push([this](int) mutable
        {
          try
          {
            // Resolve hanging reference to python thread state
            {
              pybind11::gil_scoped_acquire getGil;
              getGil.dec_ref();
            }
            
            // Pump the asyncio loop for a specified number of milliseconds, release Gil
            // and Sleep.  If the event loop has no active tasks, sleep for 4x as long.
            size_t nTasks = 0;
            while (!_stopped)
            {
              {
                pybind11::gil_scoped_acquire getGil;
                nTasks = _pumpFunction(_eventLoop, asyncioTimeout / 1000 * (nTasks > 0 ? 1 : 0.25)).cast<size_t>();
              }
              Sleep(nTasks > 0 ? sleepTime : sleepTime * 4);
            }

            // Aquire GIL to decref our python objects and close the event loop
            {
              pybind11::gil_scoped_acquire getGil;
              _eventLoop.attr("close");
              _pumpFunction     = pybind11::object();
              _callSoonFunction = pybind11::object();
              _eventLoop        = pybind11::object();
            }
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Failed to initialise python worker thread: {0}", e.what());
          }
        });
      }
      ~EventLoop()
      {
        try
        {
          if (_thread.size() > 0)
            shutdown();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Failed to shutdown python worker thread: {0}", e.what());
        }
      }
      template <class...Args>
      void callback(const pybind11::object& func, Args&&... args)
      {
        if (!active())
          return;
        auto loggedFunc = pybind11::module::import("xloil.register").attr("_logged_wrapper")(func);
        _callSoonFunction(loggedFunc, std::forward<Args>(args)...);
      }
      template <class...Args>
      void callback(const char* module, const char* func, Args&&... args)
      {
        callback(pybind11::module::import(module).attr(func), std::forward<Args>(args)...);
      }

      void runAsync(const pybind11::object& coro)
      {
        auto loggedCoro = pybind11::module::import("xloil.register").attr("_logged_wrapper_async")(coro);
        pybind11::module::import("asyncio").attr("run_coroutine_threadsafe")(loggedCoro, _eventLoop);
      }
      bool active()
      {
        return !_stopped;
      }
      void stop()
      {
        _stopped = true;
      }
      void shutdown()
      {
        stop();
        // Don't empty queue on thread stop to allow our event loop to shutdown cleanly
        _thread.stop(true);
      }

      auto& loop() { return _eventLoop; }
      auto& thread() { return _thread.get_thread(0); }
    };
  }
}