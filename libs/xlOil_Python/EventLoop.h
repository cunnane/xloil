#include <pybind11/pybind11.h>
#include <memory>
#include <thread>
#include <atomic>

namespace ctpl {
  class thread_pool;
}

namespace xloil
{
  namespace Python
  {
    class EventLoop
    {
      std::unique_ptr<ctpl::thread_pool> _thread;
      pybind11::object _eventLoop;
      pybind11::object _callSoonFunction;
      pybind11::object _pumpFunction;
      std::atomic<bool> _stopped;

      pybind11::object loggedWrapper(const pybind11::object& func);

    public:
      unsigned asyncioTimeout;
      unsigned sleepTime;

      EventLoop(unsigned asyncioTimeout_ = 200, unsigned sleepTime_ = 200);

      ~EventLoop();

      template <class...Args>
      void callback(const pybind11::object& func, Args&&... args)
      {
        if (!active())
          return;
        _callSoonFunction(loggedWrapper(func), std::forward<Args>(args)...);
      }

      template <class...Args>
      void callback(const char* module, const char* func, Args&&... args)
      {
        callback(pybind11::module::import(module).attr(func), std::forward<Args>(args)...);
      }

      void runAsync(const pybind11::object& coro);

      bool active()
      {
        return !_stopped;
      }

      void stop()
      {
        _stopped = true;
      }

      void shutdown();

      pybind11::object& loop();
      std::thread& thread();
    };
  }
}