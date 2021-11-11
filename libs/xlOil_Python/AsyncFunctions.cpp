#include "AsyncFunctions.h"
#include "PyFunctionRegister.h"
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
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
    class EventLoopController
    {
      std::atomic<bool> _stopped = true;
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
              py::gil_scoped_acquire getGil;
              getGil.inc_ref();
              // TODO: not sure calling back into xloil.register is that cool a design...
              const auto xloilMod = py::module::import("xloil.register");
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
        if (!_stopped)
          return;
        
        _loopRunner = _thread.push(
          [this](int)
          {
          _stopped = false;
            try
            {
              constexpr double timeout = 0.5; // seconds 
              bool tasks = true;
              do
              {
                py::gil_scoped_acquire getGil;
                tasks = _runLoopFunction(this->_eventLoop, timeout).cast<int>() > 0;
              } while (!_stopped && tasks);
            }
            catch (const std::exception& e)
            {
              XLO_ERROR("Error running asyncio loop: {0}", e.what());
            }
            _stopped = true;
          }
        );
      }


      bool stopped()
      {
        return _stopped;
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
        _stopped = true;
        // Must wait for the loop to stop, otherwise we may switch the stop
        // flag back without it ever being checked
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

    auto& eventLoopController()
    {
      return EventLoopController::instance();
    }


    struct AsyncReturn : public AsyncHelper
    {
      AsyncReturn(
        const ExcelObj& asyncHandle,
        const shared_ptr<const IPyToExcel>& returnConverter,
        CallerInfo&& caller)
        : AsyncHelper(asyncHandle)
        , _returnConverter(returnConverter)
        , _caller(std::move(caller))
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
        eventLoopController().start();
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
          eventLoopController().callback(_task.attr("cancel"));
          _task.release();
        }
      }

      const CallerInfo& caller() const noexcept
      {
        return _caller;
      }

    private:
      shared_ptr<const IPyToExcel> _returnConverter;
      py::object _task;
      CallerInfo _caller;
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
        vector<py::object> args(1 + info->argArraySize());

        // Raw ptr, but we take ownership in the next line
        auto* asyncReturn = new AsyncReturn(
          *asyncHandle,
          info->getReturnConverter(),
          CallerInfo());

        args[PyFuncInfo::theVectorCallOffset] = py::cast(asyncReturn,
          py::return_value_policy::take_ownership);

        py::object kwargs;
        info->convertArgs(xlArgs + 1, (PyObject**)(args.data() + 1), kwargs);

        info->invoke(args, kwargs.ptr());
      }
      catch (const py::error_already_set& e)
      {
        raiseUserException(e);
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
        const CallerInfo& caller)
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
        eventLoopController().start();
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
          result = pyCacheAdd(value, _caller.writeInternalAddress().c_str());

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
        eventLoopController().callback(_task.attr("cancel"));
      }
      bool done() noexcept
      {
        return !_running;
      }
      void wait() noexcept
      {
        // asyncio.Future has no 'wait'
      }
      const CallerInfo& caller() const noexcept
      {
        return _caller;
      }
    private:
      IRtdPublish& _notify;
      shared_ptr<const IPyToExcel> _returnConverter;
      py::object _task;
      std::atomic<bool> _running = true;
      const CallerInfo& _caller;
    };

    /// <summary>
    /// Holder for python target function and its arguments.
    /// Able to compare arguments with another AsyncTask
    /// </summary>
    struct RtdAsyncTask : public IRtdAsyncTask
    {
      const PyFuncInfo& _info;
      vector<py::object> _args;
      PyObject *_kwargs;
      shared_ptr<RtdReturn> _returnObj;
      CallerInfo _caller;

      /// <summary>
      /// Steals references to PyObjects
      /// </summary>
      RtdAsyncTask(const PyFuncInfo& info, vector<py::object>&& args, PyObject* kwargs)
        : _info(info)
        , _args(args)
        , _kwargs(kwargs)
      {}

      virtual ~RtdAsyncTask()
      {
        py::gil_scoped_acquire gilAcquired;
        _args.clear();
        Py_XDECREF(_kwargs);
        _returnObj.reset();
      }

      void start(IRtdPublish& publish) override
      {
        _returnObj.reset(new RtdReturn(publish, _info.getReturnConverter(), _caller));
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        _args[PyFuncInfo::theVectorCallOffset] = py::cast(_returnObj);
        _info.invoke(_args, _kwargs);
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

        if (_args.size() != that->_args.size())
          return false;

        // Skip first argument as that contains the the RtdReturn object which will
        // be different (set to None in unstarted tasks)
        auto nSkip = 1 + PyFuncInfo::theVectorCallOffset;
        for (auto i = _args.begin() + nSkip, j = that->_args.begin() + nSkip;
          i != _args.end();
          ++i, ++j)
        {
          if (!i->equal(*j))
            return false;
        }

        if (!_kwargs)
          return !that->_kwargs;
        
        auto kwargs = PyBorrow<py::dict>(_kwargs);
        auto that_kwargs = PyBorrow<py::dict>(that->_kwargs);
        
        if (kwargs.size() != that_kwargs.size())
          return false;

        for (auto i = kwargs.begin(); i != kwargs.end(); ++i)
        {
          if (!i->second.equal(that_kwargs[i->first]))
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
        PyObject *kwargsP;

        // Array size +1 to allow for RtdReturn argument
        vector<py::object> args(1 + info->argArraySize());
        {
          py::gil_scoped_acquire gilAcquired;

          py::object kwargs;
          // +1 to skip the RtdReturn argument
          info->convertArgs(xlArgs, (PyObject**)(args.data() + 1), kwargs);

          // TODO: not sure of the need for this
          kwargsP = kwargs.release().ptr();
        }

        auto value = rtdAsync(
          std::make_shared<RtdAsyncTask>(*info, std::move(args), kwargsP));

        return returnValue(value ? *value : CellError::NA);
      }
      catch (const py::error_already_set& e)
      {
        raiseUserException(e);
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
          .def("set_task", &AsyncReturn::set_task)
          .def("caller", &AsyncReturn::caller);

        py::class_<RtdReturn, shared_ptr<RtdReturn>>(mod, "RtdReturn")
          .def("set_result", &RtdReturn::set_result)
          .def("set_done", &RtdReturn::set_done)
          .def("set_task", &RtdReturn::set_task)
          .def("caller", &RtdReturn::caller);


        mod.def("get_event_loop", []() { return eventLoopController().getEventLoop(); });
      });

      // Uncomment this for debugging in case weird things happen with the GIL not releasing
      //static auto gilCheck = Event::AfterCalculate().bind([]() { XLO_INFO("PyGIL State: {}", PyGILState_Check());  });
    }
  }
}