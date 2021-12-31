#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyFuture.h"
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelThread.h>
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
    using ExcelObjFuture = PyFuture<ExcelObj, PyFromAny>;

    auto callExcelAsync(const py::object& func, const py::args& args)
    {
      // Convert all args to Excel objects
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

      for (auto i = 0; i < nArgs; ++i)
        xlArgs.emplace_back(FromPyObj<false, ExcelType::Missing>()(args[i].ptr()));

      py::gil_scoped_release releaseGil;

      // Run the function on the main thread
      return ExcelObjFuture(runExcelThread([funcNum, args = std::move(xlArgs)]() 
      {
        ExcelObj result;
        auto ret = xloil::callExcelRaw(funcNum, &result, args.size(), args.begin());
        if (ret != 0)
          result = std::wstring(L"#") + xlRetCodeToString(ret);
        return std::move(result);
      }, ExcelRunQueue::XLL_API));
    }

    auto callExcel(const py::object& func, const py::args& args)
    {
      return callExcelAsync(func, args).result();
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        ExcelObjFuture::bind(mod, "ExcelObjFuture");

        mod.def("run", callExcel, py::arg("func"));
        mod.def("run_async", callExcelAsync, py::arg("func"));
      });
    }
  }
}