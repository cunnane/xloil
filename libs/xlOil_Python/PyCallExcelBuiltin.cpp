#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyFuture.h"
#include "PyCore.h"
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
        else if (isRangeType(p))
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
        {
          if (ret == msxll::xlretInvXloper && funcNum == msxll::xlUDF)
            result = formatStr(L"#Unrecognised function '%s'", args[0].toString().c_str());
          else
            result = wstring(L"#") + xlRetCodeToString(ret);
        }
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

      auto funcName = to_wstring(func);

      py::gil_scoped_release releaseGil;

      return ExcelObjFuture(runExcelThread([
          funcName = std::move(funcName),
          args = std::move(xlArgs)
        ]()
        {
          const ExcelObj* argsP[30];
          for (size_t i = 0; i < args.size(); ++i)
            argsP[i] = &args[i];
          return thisApp().run(funcName, args.size(), argsP);
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
        ExcelObjFuture::bind(mod, "_ExcelObjFuture");

        mod.def("run", 
          appRun, 
          R"(
            Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
            This can call any user-defined function or macro but not built-in functions.

            The type and order of arguments expected depends on the function being called.

            Must be called on Excel's main thread, for example in worksheet function or 
            command.
          )",
          py::arg("func"));

        mod.def("run_async", 
          appRunAsync, 
          R"(
            Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
            This can call any user-defined function or macro but not built-in functions.

            Calls to the Excel API must be done on Excel's main thread: this async function
            can be called from any thread but will require the main thread to be available
            to return a result.

            Returns an **awaitable**, i.e. a future which holds the result.
          )",
          py::arg("func"));

        mod.def("call", 
          callXll, 
          R"(
            Calls a built-in worksheet function or command or a user-defined function with the 
            given name. The name is case-insensitive; built-in functions take priority in a name
            clash.
            
            The type and order of arguments expected depends on the function being called.  

            `func` can be built-in function number (as an int) which slightly reduces the lookup overhead

            This function must be called from a *non-local worksheet function on the main thread*.

            `call` can also invoke old-style `macro sheet commands <https://docs.excel-dna.net/assets/excel-c-api-excel-4-macro-reference.pdf>`_
          )",
          py::arg("func"));

        mod.def("call_async", 
          callXllAsync, 
          R"(
            Calls a built-in worksheet function or command or a user-defined function with the 
            given name.  See ``xloil.call``.

            Calls to the Excel API must be done on Excel's main thread: this async function
            can be called from any thread but will require the main thread to be available
            to return a result.
            
            Returns an **awaitable**, i.e. a future which holds the result.
          )",
          py::arg("func"));
      });
    }
  }
}