#include "EventLoop.h"
#include "PyHelpers.h"
#include "PyEvents.h"
#include <xlOil/Preprocessor.h>
#include <CTPL/ctpl_stl.h>
#include <xlOil/Log.h>
#define ARSE BOOST_PP_CAT("foo", "bar")
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
#if PY_VERSION_HEX < 0x03070000
#define ALL_TASKS "asyncio.Task.all_tasks"
#else
#define ALL_TASKS "asyncio.all_tasks"
#endif


#define XLO_PUMP_CODE BOOST_PP_CAT(BOOST_PP_CAT(R"(
def pump(loop, timeout):

    import asyncio

    async def wait():
      await asyncio.sleep(timeout)

    loop.run_until_complete(wait())

    all_tasks = )", ALL_TASKS) R"(
    return len([task for task in all_tasks(loop) if not task.done()])
)")

    // We create the loop pumper here rather than in the xloil package to avoid 
    // a dependency on that package (in particular, processing __init__. It's
    // rather difficult to define async functions in pybind so we create the 
    // necessary object with PyRun_String.
    auto _create_pump_function()
    {
      py::dict globals;
      py::dict locals;

      auto result = PyRun_String(XLO_PUMP_CODE, Py_file_input, globals.ptr(), locals.ptr());
      // Very annoying undef, not sure how to do it earlier
#undef ALL_TASKS

      if (result == nullptr)
        throw py::error_already_set();
      Py_XDECREF(result);
      // locals["pump"] doesn't work, returns a dict
      auto func = PyBorrow(PyDict_GetItemString(locals.ptr(), "pump"));
      return func;
    }


    EventLoop::EventLoop(unsigned asyncioTimeout_, unsigned sleepTime_)
      : _thread(new ctpl::thread_pool(1))
      , _stopped(false)
      , asyncioTimeout(asyncioTimeout_)
      , sleepTime(sleepTime_)
    {
#ifdef _DEBUG
      if (PyGILState_Check() == 1)
        XLO_THROW("Release GIL before constructing an EventLoop");
#endif

      _thread->push([this](int) mutable
        {
          try
          {
            pybind11::gil_scoped_acquire getGil;
            // Create a hanging ref to the python thread state, to avoid all context vars,
            // particularly the asyncio loop, being deleted at the end of this function.
            getGil.inc_ref();

            _eventLoop = pybind11::module::import("asyncio").attr("new_event_loop")();
            _pumpFunction = _create_pump_function();
            _callSoonFunction = _eventLoop.attr("call_soon_threadsafe");
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Failed to initialise python event loop: {0}", e.what());
          }
        }).get();

        if (!_eventLoop || _eventLoop.is_none() || !_callSoonFunction || _callSoonFunction.is_none())
          XLO_THROW("Failed starting event loop");

        _thread->push([this](int) mutable
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
                _pumpFunction = pybind11::object();
                _callSoonFunction = pybind11::object();
                _eventLoop = pybind11::object();
              }
            }
            catch (const std::exception& e)
            {
              XLO_ERROR("Failed to initialise python worker thread: {0}", e.what());
            }
          });
    }

    EventLoop::~EventLoop()
    {
      try
      {
        if (_thread->size() > 0)
          shutdown();
        assert(!_pumpFunction);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Failed to shutdown python worker thread: {0}", e.what());
      }
    }

    void EventLoop::runAsync(const pybind11::object& coro)
    {
      auto loggedCoro = pybind11::module::import("xloil._event_loop").attr("_logged_wrapper_async")(coro);
      pybind11::module::import("asyncio").attr("run_coroutine_threadsafe")(loggedCoro, _eventLoop);
    }

    void EventLoop::shutdown()
    {
      stop();
      // Don't empty queue on thread stop to allow our event loop to shutdown cleanly
      _thread->stop(true);
    }

    pybind11::object EventLoop::loggedWrapper(const pybind11::object& func)
    {
      return pybind11::module::import("xloil._event_loop").attr("_logged_wrapper")(func);
    }

    pybind11::object& EventLoop::loop()
    { 
      return _eventLoop; 
    }
    std::thread& EventLoop::thread()
    { 
      return _thread->get_thread(0); 
    }
  }
}