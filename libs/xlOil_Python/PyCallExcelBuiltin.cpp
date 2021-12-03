#include "PyHelpers.h"
#include "TypeConversion/BasicTypes.h"
#include "PyFuture.h"
#include <xloil/ExcelCall.h>
#include <xlOil/ExcelApp.h>
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
      // func can be a string or Excel function number
      int funcNum;
      if (PyLong_Check(func.ptr()))
        funcNum = PyLong_AsLong(func.ptr());
      else
        funcNum = excelFuncNumber(((string)py::str(func)).c_str());

      if (funcNum < 0)
        throw py::value_error("Not an Excel function: " + (string)py::str(func));

      // Convert all args to Excel objects
      auto nArgs = args.size();
      vector<ExcelObj> xlArgs;
      xlArgs.reserve(nArgs);
      for (auto i = 0; i < nArgs; ++i)
        xlArgs.emplace_back(FromPyObj<false>()(args[i].ptr()));

      py::gil_scoped_release releaseGil;

      // Run the function on the main thread
      return ExcelObjFuture(runExcelThread([funcNum, args = std::move(xlArgs)]() {
        ExcelObj result;
        auto ret = xloil::callExcelRaw(funcNum, &result, args.size(), args.begin());
        if (ret != 0)
          result = CellError::Value;
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

        mod.def("excel_func", callExcel, py::arg("func"));
        mod.def("excel_func_async", callExcelAsync, py::arg("func"));
      });
    }
  }
}