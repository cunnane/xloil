#include "AsyncFunctions.h"
#include "FunctionRegister.h"
#include "InjectedModule.h"
#include "BasicTypes.h"
#include "PyHelpers.h"
#include <xloil/ExcelObj.h>
#include <xloil/Async.h>
#include <xloil/RtdServer.h>
#include <xloil/StaticRegister.h>
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

    struct AsyncReturn : public AsyncHelper
    {
      AsyncReturn(
        const ExcelObj& asyncHandle,
        const shared_ptr<IPyToExcel>& returnConverter)
        : AsyncHelper(asyncHandle)
        , _returnConverter(returnConverter)
      {}

      ~AsyncReturn()
      {
        // This aborts, when it tries to call cancel on the task. Not sure why
        //cancel();
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

      void cancel() override
      {
        if (_task.ptr())
        {
          py::gil_scoped_acquire gilAcquired;
          _task.attr("cancel").call();
          _task.release();
        }
      }

    private:
      shared_ptr<IPyToExcel> _returnConverter;
      py::object _task;
    };

    inline ctpl::thread_pool* getWorkerThreadPool()
    {
      static auto* workerPool = []()
      {
        constexpr size_t nThreads = 1;
        auto* pool = new ctpl::thread_pool(nThreads);

        // We create a hanging reference in gil_scoped_aquire to prevent it 
        // destroying the python thread state. The thread state contains thread 
        // locals used by asyncio to find the event loop for that thread and avoid
        // creating a new one.
        pool->push([](int)
        {
          py::gil_scoped_acquire acquire;
          acquire.inc_ref();
        });
        return pool;
      }();

      static auto workerPoolDeleter = Event_PyBye().bind([]
      {
        if (workerPool)
        {
          // Resolve the hanging reference in gil_scoped_aquire and destroy
          // the python thread state
          workerPool->push([](int)
          {
            py::gil_scoped_acquire acquire;
            acquire.dec_ref();
          });
          workerPool->stop();
          delete workerPool;
        }
      });

      return workerPool;
    }

    struct EventLoopController
    {
      std::atomic<bool> _stop = true;
      PyObject* _runLoopFunction = nullptr;

      EventLoopController()
      {}

      void start()
      {
        _stop = false;

        py::gil_scoped_acquire gilAcquired;

        // This should always be run on the main thread, also the 
        // GIL gives us a mutex.
        if (!_runLoopFunction)
        {
          const auto xloilModule = py::module::import("xloil");
          _runLoopFunction = xloilModule.attr("_pump_message_loop").ptr();
        }

        getWorkerThreadPool()->push(
          [=](int)
          {
            py::gil_scoped_acquire gilAcquired;
            try
            {
              py::reinterpret_borrow<py::function>(_runLoopFunction)
                .call(&theLoopController);
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
      /// the function to run.
      /// </summary>
      template<typename F>
      auto runInterrupt(F && f)
      {
        // TODO: could call through another thread rather than stopping the event loop?
        stop();
        auto ret = getWorkerThreadPool()->push(f);
        start();
        return ret;
      }
      void stop()
      {
        _stop = true;
      }
      bool stopped()
      {
        return _stop;
      }
    } theLoopController;

    void pythonAsyncCallback(
      PyFuncInfo* info,
      const ExcelObj* asyncHandle,
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        PyObject *argsP, *kwargsP;
        AsyncReturn* asyncReturn;

        {
          py::gil_scoped_acquire gilAcquired;

          PyErr_Clear();

          // I think it's better to process the arguments to python here rather than 
          // copying the ExcelObj's and converting on the async thread (since CPython
          // is single threaded anyway)
          auto[args, kwargs] = info->convertArgs(xlArgs);
          if (kwargs.is_none())
            kwargs = py::dict();

          // Raw ptr, but we take ownership below
          asyncReturn = new AsyncReturn(
            *asyncHandle,
            info->returnConverter);

          // Kwargs will be dec-refed after the asyncReturn is used
          kwargs[THREAD_CONTEXT_TAG] = py::cast(asyncReturn,
            py::return_value_policy::take_ownership);

          // Need to drop pybind links before lambda capture in otherwise the lambda's 
          // dtor is called at some random time after losing the GIL and it crashes.
          argsP = args.release().ptr();
          kwargsP = kwargs.release().ptr();
        }

        auto functor = [info, argsP, kwargsP, asyncReturn](int threadId) mutable
        {
          py::gil_scoped_acquire gilAcquired;
          {
            try
            {
              // This will return through the asyncReturn object
              info->invoke(argsP, kwargsP);
            }
            catch (const std::exception& e)
            {
              asyncReturn->result(ExcelObj(e.what()));
            }
            Py_XDECREF(argsP);
            Py_XDECREF(kwargsP);
          }
        };
        theLoopController.runInterrupt(functor);
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
        const shared_ptr<IPyToExcel>& returnConverter)
        : _notify(notify)
        , _returnConverter(returnConverter)
      {}
      ~RtdReturn()
      {
        // TODO: maybe use a raw PyObject to avoid needing this?
        if (_task.ptr())
        {
          py::gil_scoped_acquire gilAcquired;
          _task.release();
        }
      }
      void set_task(const py::object& task)
      {
        _task = task;
      }
      void set_result(const py::object& value)
      {
        py::gil_scoped_acquire gilAcquired;
        ExcelObj result = _returnConverter
          ? (*_returnConverter)(*value.ptr())
          : FromPyObj()(value.ptr());
        _notify.publish(std::move(result));
      }

      void cancel()
      {
        if (_task.ptr())
        {
          py::gil_scoped_acquire gilAcquired;
          _task.attr("cancel").call();
          _task.release();
        }
      }

    private:
      IRtdPublish& _notify;
      shared_ptr<IPyToExcel> _returnConverter;
      py::object _task;
    };

    /// <summary>
    /// Holder for python target function and its arguments.
    /// Able to compare arguments with another AsyncTask
    /// </summary>
    struct AsyncTask : public IRtdAsyncTask
    {
      PyFuncInfo* _info;
      PyObject *_args, *_kwargs;
      RtdReturn* _returnObj = nullptr;
      std::future<void> _future;

      /// <summary>
      /// Steals references to PyObjects
      /// </summary>
      AsyncTask(PyFuncInfo* info, PyObject* args, PyObject* kwargs)
        : _info(info)
        , _args(args)
        , _kwargs(kwargs)
      {}

      virtual ~AsyncTask()
      {
        py::gil_scoped_acquire gilAcquired;
        Py_XDECREF(_args);
        Py_XDECREF(_kwargs);
        delete _returnObj;
      }

      void start(IRtdPublish& publish) override
      {
        _returnObj = new RtdReturn(publish, _info->returnConverter);

        _future = theLoopController.runInterrupt(
          [=](int /*threadId*/)
          {
            py::gil_scoped_acquire gilAcquired;

            PyErr_Clear();

            auto kwargs = PyBorrow<py::object>(_kwargs);
            kwargs[THREAD_CONTEXT_TAG] = _returnObj;

            _info->invoke(_args, kwargs.ptr());
          }
        );
      }
      bool done() override
      {
        return !_future.valid()
          || _future.wait_for(std::chrono::seconds(0)) == std::future_status::ready;
      }
      void wait() override
      {
        if (_future.valid())
          _future.wait();
      }
      void cancel() override
      {
        if (_returnObj)
          _returnObj->cancel();
      }
      bool operator==(const IRtdAsyncTask& that_) const override
      {
        py::gil_scoped_acquire gilAcquired;

        const auto& that = (const AsyncTask&)that_;
        auto args = PyBorrow<py::tuple>(_args);
        auto kwargs = PyBorrow<py::dict>(_kwargs);
        auto that_args = PyBorrow<py::tuple>(that._args);
        auto that_kwargs = PyBorrow<py::dict>(that._kwargs);

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
      PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept
    {
      try
      {
        // TODO: consider argument capture and equality check under c++
        PyObject *argsP, *kwargsP;
        {
          py::gil_scoped_acquire gilAcquired;

          auto[args, kwargs] = info->convertArgs(xlArgs);
          if (kwargs.is_none())
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
          std::make_shared<AsyncTask>(info, argsP, kwargsP));
        return returnValue(value ? *value : CellError::NA);
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
          .def("set_task", &AsyncReturn::set_task);

        py::class_<RtdReturn>(mod, "RtdReturn")
          .def("set_result", &RtdReturn::set_result)
          .def("set_task", &RtdReturn::set_task);

        py::class_<EventLoopController>(mod, "EventLoopController")
          .def("stopped", &EventLoopController::stopped);

        // This is a module level string so async function can find the
        // control object we add to their keyword args.
        mod.add_object("ASYNC_CONTEXT_TAG", py::str(THREAD_CONTEXT_TAG));
      });
    }
  }
}