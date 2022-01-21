#include "PyHelpers.h"

namespace py = pybind11;
using std::wstring;

namespace pybind11
{
  namespace
  {
    static std::mutex theTracebackLock;
  }
  std::string error_full_traceback()
  {
    if (!PyErr_Occurred())
    {
      PyErr_SetString(PyExc_RuntimeError, "Attempt to throw python error without indicator set");
      return "Unknown internal error occurred";
    }

    // We store the error indicator and restore it on exit. This allows the
    // ctor of error_already_set to grab the indicator using PyErr_Fetch.
    error_scope error;

    // Python's error output from PyErr_Print is backwards, so we output the
    // original error first at that's likely the most useful thing to see in the
    // cell where the result is shown
    std::string errorString;
    if (error.type)
    {
      errorString += handle(error.type).attr("__name__").cast<std::string>();
      errorString += ": ";
    }
    if (error.value)
    {
      errorString += (std::string)str(error.value);
      errorString += "\n";
    }

    // Python only provides a facility for writing an error to stderr via
    // PyErr_Print. So we replace stderr with a StringIO stream
    auto ioMod = PyImport_ImportModule("io");
    auto stringIO = PyObject_CallMethod(ioMod, "StringIO", NULL);
    Py_DECREF(ioMod);

    // Protect the change to stderr which is global - we only want one 
    // thread trying this trick at a time!
    std::scoped_lock lock(theTracebackLock);

    auto previousStdErr = PySys_GetObject("stderr");
    PySys_SetObject("stderr", stringIO);

    // Restore the error and call PyErr_Print which clears the error indicator.
    // The dtor of error_scope will restore it again on exit from this function.
    if (error.type) Py_INCREF(error.type);
    if (error.value) Py_INCREF(error.value);
    if (error.trace) Py_INCREF(error.trace);
    PyErr_Restore(error.type, error.value, error.trace);
    PyErr_Print();

    PySys_SetObject("stderr", previousStdErr);
    Py_DECREF(previousStdErr);

    // Grab the string output from stringIO and cleanup 
    auto fullTrace = PyObject_CallMethod(stringIO, "getvalue", NULL);
    errorString += (std::string)str(fullTrace);
    Py_DECREF(stringIO);
    Py_DECREF(fullTrace);

    return errorString;
  }
}
namespace xloil {
  namespace Python
  {
    bool getItemIndexReader1d(
      const pybind11::object& index,
      const size_t size, size_t& from, size_t& to)
    {
      if (index.is_none())
      {
        from = 0;
        to = size;
      }
      else if (PySlice_Check(index.ptr()))
      {
        size_t sliceLength, step;
        index.cast<pybind11::slice>().compute(size, &from, &to, &step, &sliceLength);
        if (step != 1)
          XLO_THROW("Slice step size must be 1");
      }
      else
      {
        from = index.cast<size_t>();
        to = from + 1;
        return true;
      }
      return false;
    }

    bool getItemIndexReader2d(
      const pybind11::tuple& loc,
      const size_t nRows, const size_t nCols,
      size_t& fromRow, size_t& fromCol,
      size_t& toRow, size_t& toCol)
    {
      if (loc.size() != 2)
        XLO_THROW("Expecting tuple of size 2");

      auto singleElement = getItemIndexReader1d(loc[0], nRows, fromRow, toRow);
      singleElement &= getItemIndexReader1d(loc[1], nCols, fromCol, toCol);
      return singleElement;
    }

    std::wstring pyToWStr(const PyObject* p)
    {
      Py_ssize_t len;
      wchar_t* wstr;
      if (!p)
        return wstring();
      else if (PyUnicode_Check(p))
        wstr = PyUnicode_AsWideCharString((PyObject*)p, &len);
      else
      {
        auto str = PyObject_Str((PyObject*)p);
        wstr = PyUnicode_AsWideCharString(str, &len);
        Py_XDECREF(str);
      }

      auto freer = std::unique_ptr<wchar_t, void(*)(void*)>(wstr, PyMem_Free);
      return wstring(wstr, len);
    }

    
  }
}