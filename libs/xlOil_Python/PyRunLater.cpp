#include "PyHelpers.h"
#include "PyFuture.h"
#include "PyCore.h"
#include <xlOil/ExcelThread.h>
#include <xlOil/StringUtils.h>


using std::shared_ptr;
using std::vector;
using std::wstring;
using std::string;
namespace py = pybind11;
using std::make_pair;

namespace xloil
{
  namespace Python
  {
    PyObject* comBusyException;

    namespace
    {
      // TODO: define in another cpp
      auto runLater(
        const py::object& callable,
        const unsigned delay,
        const unsigned retryPause,
        wstring& api)
      {
        int flags = 0;
        toLower(api);
        if (api.empty() || api.find(L"com") != wstring::npos)
          flags |= ExcelRunQueue::COM_API;
        if (api.find(L"xll") != wstring::npos)
          flags |= ExcelRunQueue::XLL_API;
        if (retryPause == 0)
          flags |= ExcelRunQueue::NO_RETRY;
        return PyFuture<PyObject*>(runExcelThread([callable = PyObjectHolder(callable)]()
          {
            py::gil_scoped_acquire getGil;
            try
            {
              return callable().release().ptr();
            }
            catch (py::error_already_set& err)
            {
              if (err.matches(comBusyException))
                throw ComBusyException();
              throw;
            }
          },
          flags,
          delay,
          retryPause));
      }

      static int theBinder = addBinder([](py::module& mod)
      {
        comBusyException = py::register_exception<ComBusyException>(mod, "ComBusyError").ptr();

        mod.def("excel_callback",
          &runLater,
          R"(
            Schedules a callback to be run in the main thread. Much of the COM API in unavailable
            during the calc cycle, in particular anything which involves writing to the sheet.
            COM is also unavailable whilst xlOil is loading.
 
            Returns a future which can be awaited.

            Parameters
            ----------

            func: callable
            A callable which takes no arguments and returns nothing

            retry : int
            Millisecond delay between retries if Excel's COM API is busy, e.g. a dialog box
            is open or it is running a calc cycle.If zero, does no retry

            wait : int
            Number of milliseconds to wait before first attempting to run this function

            api : str
            Specify 'xll' or 'com' or both to indicate which APIs the call requires.
            The default is 'com': 'xll' would only be required in rare cases.
            )",
          py::arg("func"),
          py::arg("wait") = 0,
          py::arg("retry") = 500,
          py::arg("api") = "");
      });
    }
  }
}