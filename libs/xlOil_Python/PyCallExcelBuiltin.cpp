#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyFuture.h"
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/AppObjects.h>
#include <future>

using std::shared_ptr;
using std::vector;
using std::wstring;
using std::string;

namespace py = pybind11;

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// Wraps the usual FromPyObj but converts None to Missing, which seems
    /// more useful in the context and Range to ExcelRef which is necessary to
    /// call many of the macro sheet commands.
    /// </summary>
    struct ArgFromPyObj
    {
      auto operator()(const py::object& obj) const
      {
        auto p = (PyObject*)obj.ptr();
        if (p == Py_None)
        {
          return ExcelObj(ExcelType::Missing);
        }
        else if (Py_TYPE(p) == rangeType)
        {
          auto* range = obj.cast<Range*>();
          return ExcelObj(refFromRange(*range));
        }
        else
          return FromPyObj<false>()(p);
      }
    };

    using ExcelObjFuture = PyFuture<ExcelObj, PyFromAny>;

    auto callXllAsync(const py::object& func, const py::args& args)
    {
      // Space to convert all args to Excel objects
      auto nArgs = args.size();
      vector<ExcelObj> xlArgs;
      xlArgs.reserve(nArgs);

      // func can be a string or Excel function number
      int funcNum;
      if (PyLong_Check(func.ptr()))
      {
        funcNum = PyLong_AsLong(func.ptr());
        if (funcNum < 0)
          throw py::value_error("Not an Excel function: " + std::to_string(funcNum));
      }
      else
      {
        const auto funcName = (string)py::str(func);
        funcNum = excelFuncNumber(funcName.c_str());
        // If we don't recognise the function name as as built-in, we try
        // to run a UDF.
        if (funcNum < 0)
        {
          funcNum = msxll::xlUDF;
          xlArgs.insert(xlArgs.begin(), ExcelObj(funcName));
        }
      }

      // Convert args with None->Missing Arg and Range->ExcelRef
      for (auto i = 0u; i < nArgs; ++i)
        xlArgs.emplace_back(ArgFromPyObj()(args[i]));

      py::gil_scoped_release releaseGil;

      // Run the function on the main thread
      return ExcelObjFuture(runExcelThread([funcNum, args = std::move(xlArgs)]()
      {
        ExcelObj result;
        auto ret = xloil::callExcelRaw(funcNum, &result, args.size(), args.begin());
        if (ret != 0)
          result = wstring(L"#") + xlRetCodeToString(ret);
        return std::move(result);
      }, ExcelRunQueue::XLL_API));
    }

    auto callXll(const py::object& func, const py::args& args)
    {
      return callXllAsync(func, args).result();
    }

    auto appRunAsync(const py::object& func, const py::args& args)
    {
      // Convert all args to Excel objects
      auto nArgs = args.size();
      if (nArgs > 30)
        throw py::value_error();

      vector<ExcelObj> xlArgs;
      xlArgs.reserve(nArgs);

      // Convert args with None->Missing Arg and Range->ExcelRef
      for (auto i = 0u; i < nArgs; ++i)
        xlArgs.emplace_back(ArgFromPyObj()(args[i]));

      py::gil_scoped_release releaseGil;

      return ExcelObjFuture(runExcelThread([
          func = pyToWStr(func), 
          args = std::move(xlArgs)
        ]()
        {
          const ExcelObj* argsP[30];
          for (size_t i = 0; i < args.size(); ++i)
            argsP[i] = &args[i];
          return excelApp().Run(func, args.size(), argsP);
        }));
    }

    auto appRun(const py::object& func, const py::args& args)
    {
      return appRunAsync(func, args).result();
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        ExcelObjFuture::bind(mod, "ExcelObjFuture");

        mod.def("run", appRun, py::arg("func"));
        mod.def("run_async", appRunAsync, py::arg("func"));
        mod.def("call", callXll, py::arg("func"));
        mod.def("call_async", callXllAsync, py::arg("func"));
      });
    }
  }
}