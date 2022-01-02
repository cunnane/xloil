#include "AsyncFunctions.h"
#include "PyFunctionRegister.h"
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyHelpers.h"
#include "PyEvents.h"
#include "EventLoop.h"
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
    auto& asyncEventLoop()
    {
      return *theCoreAddin().thread;
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
      {
      }

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
          if (py::hasattr(_task, "cancel"))
            asyncEventLoop().callback(_task.attr("cancel"));
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
      {
      }
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
        asyncEventLoop().callback(_task.attr("cancel"));
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
      py::object _kwargs;
      shared_ptr<RtdReturn> _returnObj;
      CallerInfo _caller;

      /// <summary>
      /// Steals references to PyObjects
      /// </summary>
      RtdAsyncTask(const PyFuncInfo& info, vector<py::object>&& args, py::object&& kwargs)
        : _info(info)
        , _args(args)
        , _kwargs(kwargs)
      {}

      virtual ~RtdAsyncTask()
      {
        py::gil_scoped_acquire gilAcquired;
        _args.clear();
        _kwargs = py::none();
        _returnObj.reset();
      }

      void start(IRtdPublish& publish) override
      {
        _returnObj.reset(new RtdReturn(publish, _info.getReturnConverter(), _caller));
        py::gil_scoped_acquire gilAcquired;

        PyErr_Clear();

        _args[PyFuncInfo::theVectorCallOffset] = py::cast(_returnObj);
        _info.invoke(_args, _kwargs.ptr());
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
        
        auto kwargs = py::dict(_kwargs);
        auto that_kwargs = py::dict(that->_kwargs);
        
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
        py::object kwargs;

        // Array size +1 to allow for RtdReturn argument
        vector<py::object> args(1 + info->argArraySize());
        {
          py::gil_scoped_acquire gilAcquired;
          
          // +1 to skip the RtdReturn argument
          info->convertArgs(xlArgs, (PyObject**)(args.data() + 1), kwargs);
        }

        // Moving a py::object means we don't need the GIL
        auto value = rtdAsync(
          std::make_shared<RtdAsyncTask>(*info, std::move(args), std::move(kwargs)));

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
          .def_property_readonly("caller", &AsyncReturn::caller)
          .def_property_readonly("loop", [](py::object x) { return asyncEventLoop().loop(); });

        py::class_<RtdReturn, shared_ptr<RtdReturn>>(mod, "RtdReturn")
          .def("set_result", &RtdReturn::set_result)
          .def("set_done", &RtdReturn::set_done)
          .def("set_task", &RtdReturn::set_task)
          .def_property_readonly("caller", &RtdReturn::caller)
          .def_property_readonly("loop", [](py::object x) { return asyncEventLoop().loop(); });


        mod.def("get_async_loop", []() { return asyncEventLoop().loop(); });
      });

      // Uncomment this for debugging in case weird things happen with the GIL not releasing
      //static auto gilCheck = Event::AfterCalculate().bind([]() { XLO_INFO("PyGIL State: {}", PyGILState_Check());  });
    }
  }
}