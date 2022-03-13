#include "PyHelpers.h"

namespace py = pybind11;
using std::wstring;

namespace pybind11
{
  // Possible faster implementation using _PyErr_Display which is a hidden API function
  // in Py 3.8+
  //   auto stringIO = module::import("io").attr("StringIO")();
  //   _PyErr_Display(stringIO, error.type, error.value, error.trace);
  //   result = handle(error.type).attr("__name__") + ": " + (std::string)str(error.value)
  //   result += stringIO("getvalue");
  // 
  std::string error_full_traceback()
  {
    if (!PyErr_Occurred())
    {
      PyErr_SetString(PyExc_RuntimeError, "Attempt to throw python error without indicator set");
      return "Unknown internal error occurred";
    }

    // Store the error indicator and restore it on exit. This allows the
    // ctor of pybind11::error_already_set to grab the indicator using PyErr_Fetch.
    error_scope error;

    // Ensures calls to the traceback module succeed
    PyErr_NormalizeException(&error.type, &error.value, &error.trace);
    
    // format_exception produces a list of strings
    auto errs = list(module::import("traceback").attr("format_exception")(
      handle(error.type), handle(error.value), handle(error.trace)));

    // Python's error output is backwards, so we show the original error first 
    // at that's likely the most useful thing to see in the cell 
    auto errorString = (std::string)str(errs[errs.size() - 1]);

    for (auto i = 0; i < errs.size(); ++i)
      errorString += (std::string)str(errs[i]);

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

    PyObject* fastCall(
      PyObject* func, PyObject* const* args, size_t nArgs, PyObject* kwargs) noexcept
    {
#if PY_VERSION_HEX < 0x03080000
      auto argTuple = PyTuple_New(nArgs);
      for (auto i = 0u; i < nArgs; ++i)
        PyTuple_SET_ITEM(argTuple, i, args[i]);

      auto retVal = kwargs
        ? PyObject_Call(func, argTuple, kwargs)
        : PyObject_CallObject(func, argTuple);

      Py_XDECREF(argTuple);
#else
      auto retVal = _PyObject_FastCallDict(
        func, args, nArgs | PY_VECTORCALL_ARGUMENTS_OFFSET, kwargs);
#endif
      return retVal;
    }
  }
}