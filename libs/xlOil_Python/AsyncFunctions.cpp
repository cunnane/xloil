#include "AsyncFunctions.h"
#include "FunctionRegister.h"
#include "PyCoreModule.h"
#include "BasicTypes.h"
#include "PyHelpers.h"
#include "PyEvents.h"
#include <xloil/ExcelObj.h>
#include <xloil/Async.h>
#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
#include <xloil/Caller.h>
#include <CTPL/ctpl_stl.h>
#include <vector>

using std::shared_ptr;
using std::vector;
using std::pair;
using std::wstring;
using std::string;
using std::make_shared;
using std::make_pair;
namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    constexpr const char* THREAD_CONTEXT_TAG = "xloil_thread_context";

    class EventLoopController
    {
      std::atomic<bool> _stop = true;
      std::future<void> _loopRunner;
      ctpl::thread_pool _thread;
      std::shared_ptr<const void> _shutdownHandler;
      py::object _eventLoop;
      py::object _runLoopFunction;
      py::object _callSoonFunction;

      EventLoopController()
        : _thread(1)
      {
        py::gil_scoped_release releaseGil;

        // We create a hanging reference in gil_scoped_aquire to prevent it 
        // destroying the python thread state. The thread state contains thread 
        // locals used by asyncio to find the event loop for that thread and avoid
        // creating a new one.
        _thread.push([self = this](int) mutable
          {
            try
            {
              py::gil_scoped_acquire acquire;
              acquire.inc_ref();
              const auto xloilMod = py::module::import("xloil.xloil");
              self->_eventLoop = xloilMod.attr("_create_event_loop")();
              self->_runLoopFunction = xloilMod.attr("_pump_message_loop");
              self->_callSoonFunction = self->_eventLoop.attr("call_soon_threadsafe");
            }
            catch (const std::exception& e)
            {
              XLO_ERROR("Failed to initialise python worker thread: {0}", e.what());
            }
          }
        ).wait();

        if (!_runLoopFunction.ptr())
          XLO_THROW("Cannot start python worker thread");

        _shutdownHandler = std::static_pointer_cast<const void>(
          Event_PyBye().bind([self = this]
          {
            self->shutdown();
          })
        );
        
        start();
      }

      EventLoopController(const EventLoopController&) = delete;
      EventLoopController& operator=(const EventLoopController&) = delete;

    public:
      void start()
      {
        _stop = false;
        _loopRunner = _thread.push(
          [=](int)
          {
            try
            {
              py::gil_scoped_acquire gilAcquired;
              // TODO: this is a bit of a busy-wait can we signal when there are tasks?
              while (!_stop)
              {
                constexpr double timeout = 0.5; // seconds 
                _runLoopFunction(this->_eventLoop, timeout);
              }
            }
            catch (const std::exception& e)
            {
              XLO_ERROR("Error running asyncio loop: {0}", e.what());
            }
          }
        );
      }
      /// <summary>
      /// Executes the function on the python asyncio thread. The thread is normally
      /// running the asyncio event loop.  This stops and restarts the loop to allow
      /// the function to run. Note we never actually use it...
      /// </summary>
      template<typename F>
      void runInterrupt(F && f)
      {
        py::gil_scoped_release noGil;
        auto task = _thread.push(f);
        stop();
        start();
      }

      bool stopped()
      {
        return _stop;
      }

      py::object getEventLoop() const
      {
        return _eventLoop;
      }

      void callback(const py::object& func)
      {
        if (!active())
          return;
        _callSoonFunction(func);
      }
      static EventLoopController& instance()
      {
        static EventLoopController instance;
        return instance;
      }
      
    private:
      void stop()
      {
        _stop = true;
        // Must wait for the loop to stop, otherwise we may switch the stop
        // flag back without it ever being check
        if (_loopRunner.valid())
          _loopRunner.wait();
      }

      void shutdown()
      {
        stop();
        // Resolve the hanging reference in gil_scoped_aquire and destroy
        // the python thread state
        _thread.push([](int)
        {
          py::gil_scoped_acquire acquire;
          acquire.dec_ref();
        }).wait();

        _thread.stop();

        py::gil_scoped_acquire acquire;
        _eventLoop = py::object();
        _runLoopFunction = py::object();
        _callSoonFunction = py::object();
      }

      bool active()
      {
        return _thread.size() > 0;
      }
    };

    auto& getLoopController()
    {
      return EventLoopController::instance();
    }


    struct AsyncReturn : public AsyncHelper
    {
      AsyncReturn(
        const ExcelObj& asyncHandle,
        const shared_ptr<const IPyToExcel>& returnConverter)
        : AsyncHelper(asyncHandle)
        , _returnConverter(returnConverter)
      {}

      ~AsyncReturn()
      {
        try
        {
          cancel();
        }
        catch (...) {}
      }

      void set_task(const py::object& task)
      {
        _task = task;
      }

      void set_result(const py::object& value)
      {
        py::gil_scoped_acquire gil;
        static ExcelObj obj = _returnConverter
          ? (*_returnConverter)(*value.ptr())
          : FromPyObj()(value.ptr());
        result(obj);
      }
      void set_done()
      {}
      
      void cancel() override
      {
        if (_task.ptr())
        {
          py::gil_scoped_acquire gilAcquired;
          getLoopController().callback(_task.attr("cancel"));
          _task.release();
        }
      }

    private:
      shared_ptr<const IPyToExcel> _returnConverter;
      py::object _task;
    };

    void pythonAsyncCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      const ExcelObj* asyncHandle = xlArgs[0];

      try
      {
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        // I think it's better to process the arguments to python here rather than 
        // copying the ExcelObj's and converting on the async thread (since CPython
        // is single threaded anyway)
        auto[args, kwargs] = info->convertArgs(xlArgs + 1);
        if (!kwargs || kwargs.is_none())
          kwargs = py::dict();

        // Raw ptr, but we take ownership below
        auto* asyncReturn = new AsyncReturn(
          *asyncHandle,
          info->getReturnConverter());

        // Kwargs will be dec-refed after the asyncReturn is used
        kwargs[THREAD_CONTEXT_TAG] = py::cast(asyncReturn,
          py::return_value_policy::take_ownership);

        info->invoke(args.ptr(), kwargs.ptr());
      }
      catch (const py::error_already_set& e) 
      {
        Event_PyUserException().fire(e.type(), e.value(), e.trace());
        asyncReturn(*asyncHandle, ExcelObj(e.what()));
      }
      catch (const std::exception& e)
      {
        XLO_WARN(e.what());
        asyncReturn(*asyncHandle, ExcelObj(e.what()));
      }
      catch (...)
      {
        XLO_WARN("Async unknown error");
        asyncReturn(*asyncHandle, ExcelObj(CellError::Value));
      }
    }

    struct RtdReturn
    {
      RtdReturn(
        IRtdPublish& notify,
        const shared_ptr<const IPyToExcel>& returnConverter,
        const wchar_t* caller)
        : _notify(notify)
        , _returnConverter(returnConverter)
        , _caller(caller)
      {}
      ~RtdReturn()
      {
        if (!_running && !_task.ptr())
          return;

        py::gil_scoped_acquire gilAcquired;
        _running = false;
        _task = py::object();
      }
      void set_task(const py::object& task)
      {
        py::gil_scoped_acquire gilAcquired;
        _task = task;
        _running = true;
      }
      void set_result(const py::object& value) const
      {
        if (!_running)
          return;
        py::gil_scoped_acquire gilAcquired;

        // Convert result to ExcelObj
        ExcelObj result = _returnConverter
          ? (*_returnConverter)(*value.ptr())
          : FromPyObj<false>()(value.ptr());

        // If nil, conversion wasn't possible, so use the cache
        if (result.isType(ExcelType::Nil))
          result = pyCacheAdd(value, _caller);

        _notify.publish(std::move(result));
      }
      void set_done()
      {
        if (!_running)
          return;
        py::gil_scoped_acquire gilAcquired;
        _running = false;
        _task = py::object();
      }
      void cancel()
      {
        if (!_running)
          return;
        py::gil_scoped_acquire gilAcquired;
        _running = false;
        getLoopController().callback(_task.attr("cancel"));
      }
      bool done() noexcept
      {
        return !_running;
      }
      void wait() noexcept
      {
        // asyncio.Future has no 'wait'
      }

    private:
      IRtdPublish& _notify;
      shared_ptr<const IPyToExcel> _returnConverter;
      py::object _task;
      std::atomic<bool> _running = true;
      const wchar_t* _caller;
    };

    /// <summary>
    /// Holder for python target function and its arguments.
    /// Able to compare arguments with another AsyncTask
    /// </summary>
    struct RtdAsyncTask : public IRtdAsyncTask
    {
      const PyFuncInfo& _info;
      PyObject *_args, *_kwargs;
      shared_ptr<RtdReturn> _returnObj;
      wstring _caller;

      /// <summary>
      /// Steals references to PyObjects
      /// </summary>
      RtdAsyncTask(const PyFuncInfo& info, PyObject* args, PyObject* kwargs)
        : _info(info)
        , _args(args)
        , _kwargs(kwargs)
        , _caller(CallerLite().writeInternalAddress())
      {}

      virtual ~RtdAsyncTask()
      {
        py::gil_scoped_acquire gilAcquired;
        Py_XDECREF(_args);
        Py_XDECREF(_kwargs);
        _returnObj.reset();
      }

      void start(IRtdPublish& publish) override
      {
        _returnObj.reset(new RtdReturn(publish, _info.getReturnConverter(), _caller.c_str()));
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        auto kwargs = PyBorrow<py::object>(_kwargs);
        kwargs[THREAD_CONTEXT_TAG] = _returnObj;

        _info.invoke(_args, kwargs.ptr());
      }
      bool done() noexcept override
      {
        return _returnObj ? _returnObj->done() : false;
      }
      void wait() noexcept override
      {
        if (_returnObj)
          _returnObj->wait();
      }
      void cancel() override
      {
        if (_returnObj)
          _returnObj->cancel();
      }
      bool operator==(const IRtdAsyncTask& that_) const override
      {
        const auto* that = dynamic_cast<const RtdAsyncTask*>(&that_);
        if (!that)
          return false;

        py::gil_scoped_acquire gilAcquired;

        auto args = PyBorrow<py::tuple>(_args);
        auto kwargs = PyBorrow<py::dict>(_kwargs);
        auto that_args = PyBorrow<py::tuple>(that->_args);
        auto that_kwargs = PyBorrow<py::dict>(that->_kwargs);

        if (args.size() != that_args.size()
          || kwargs.size() != that_kwargs.size())
          return false;

        for (auto i = args.begin(), j = that_args.begin();
          i != args.end();
          ++i, ++j)
        {
          if (!i->equal(*j))
            return false;
        }
        for (auto i = kwargs.begin(); i != kwargs.end(); ++i)
        {
          if (!i->first.equal(py::str(THREAD_CONTEXT_TAG))
            && !i->second.equal(that_kwargs[i->first]))
            return false;
        }
        return true;
      }
    };

    ExcelObj* pythonRtdCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        // TODO: consider argument capture and equality check under c++
        PyObject *argsP, *kwargsP;
        {
          py::gil_scoped_acquire gilAcquired;

          auto[args, kwargs] = info->convertArgs(xlArgs);
          if (!kwargs || kwargs.is_none())
            kwargs = py::dict();

          // Add this here so that dict sizes for running and newly 
          // created tasks match
          kwargs[THREAD_CONTEXT_TAG] = py::none();

          // Need to drop pybind links before capturing in lambda otherwise the destructor
          // is called at some random time after losing the GIL and it crashes.
          argsP = args.release().ptr();
          kwargsP = kwargs.release().ptr();
        }

        auto value = rtdAsync(
          std::make_shared<RtdAsyncTask>(*info, argsP, kwargsP));

        return returnValue(value ? *value : CellError::NA);
      }
      catch (const py::error_already_set& e)
      {
        Event_PyUserException().fire(e.type(), e.value(), e.trace());
        return returnValue(e.what());
      }
      catch (const std::exception& e)
      {
        return returnValue(e.what());
      }
      catch (...)
      {
        return returnValue(CellError::Null);
      }
    }
    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<AsyncReturn>(mod, "AsyncReturn")
          .def("set_result", &AsyncReturn::set_result)
          .def("set_done", &AsyncReturn::set_done)
          .def("set_task", &AsyncReturn::set_task);

        py::class_<RtdReturn, shared_ptr<RtdReturn>>(mod, "RtdReturn")
          .def("set_result", &RtdReturn::set_result)
          .def("set_done", &RtdReturn::set_done)
          .def("set_task", &RtdReturn::set_task);


        mod.def("get_event_loop", []() { return getLoopController().getEventLoop(); });
      });

      // Uncomment this for debugging in case weird things happen with the GIL not releasing
      //static auto gilCheck = Event::AfterCalculate().bind([]() { XLO_INFO("PyGIL State: {}", PyGILState_Check());  });
    }
  }
}