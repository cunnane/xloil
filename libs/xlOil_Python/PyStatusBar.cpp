#include "PyCore.h"
#include <xlOil/ExcelUI.h>

using std::shared_ptr;
using std::wstring_view;
using std::vector;
using std::wstring;
namespace py = pybind11;


namespace xloil
{
  namespace Python
  {
    namespace
    {
      struct PyStatusBar
      {
        std::unique_ptr<StatusBar> _bar;

        PyStatusBar(size_t timeout)
          : _bar(new StatusBar(timeout))
        {}
        void msg(const std::wstring& msg, size_t timeout)
        {
          py::gil_scoped_release releaseGil;
          _bar->msg(msg, timeout);
        }
        void exit(py::args)
        {
          py::gil_scoped_release releaseGil;
          _bar.reset();
        }
      };


      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyStatusBar>(mod, "StatusBar", R"(             
            Displays status bar messages and clears the status bar (after an optional delay) 
            on context exit.

            Examples
            --------

            ::

              with StatusBar(1000) as status:
                status.msg('Doing slow thing')
                ...
                status.msg('Done slow thing')
          )")
          .def(py::init<size_t>(), 
              R"(
                Constructs a StatusBar with a timeout specified in milliseconds.  After the 
                StatusBar context exits, any messages will be cleared after the timeout
              )",
              py::arg("timeout") = 0)
          .def("__enter__", [](py::object self) { return self; })
          .def("__exit__", &PyStatusBar::exit)
          .def("msg", &PyStatusBar::msg,
            R"(
              Posts a status bar message, and if `timeout` is non-zero, clears if after
              the specified number of milliseconds
            )",
            py::arg("msg"), py::arg("timeout") = 0);
      });
    }
  }
}