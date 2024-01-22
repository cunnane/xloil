#include "PyHelpers.h"

namespace py = pybind11;
using std::wstring;


std::wstring to_wstring(const PyObject* p)
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
    if (!str)
      throw py::error_already_set();
    wstr = PyUnicode_AsWideCharString(str, &len);
    Py_XDECREF(str);
  }

  if (!wstr)
    throw py::error_already_set();

  auto freer = std::unique_ptr<wchar_t, void(*)(void*)>(wstr, PyMem_Free);
  return wstring(wstr, len);
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
        from = PyLong_AsLong(index.ptr());
        if (from == -1 && PyErr_Occurred())
          XLO_THROW("Could not read index as a slice or int");
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

    PyObject* fastCall(
      PyObject* func, PyObject* const* args, size_t nArgs, PyObject* kwargs) noexcept
    {
#if PY_VERSION_HEX < 0x03080000
      auto argTuple = PyTuple_New(nArgs);
      if (!argTuple)
        return nullptr;

      for (auto i = 0u; i < nArgs; ++i)
      {
        PyTuple_SET_ITEM(argTuple, i, args[i]);
        Py_XINCREF(args[i]);
      }

      auto retVal = PyObject_Call(func, argTuple, kwargs);

      Py_XDECREF(argTuple);
#else
      
#if PY_VERSION_HEX < 0x03090000
#  define PyObject_VectorcallDict _PyObject_FastCallDict 
#  define PyObject_Vectorcall _PyObject_Vectorcall 
#endif

      auto retVal = kwargs
        ? PyObject_VectorcallDict(func, args, nArgs | PY_VECTORCALL_ARGUMENTS_OFFSET, kwargs)
        : PyObject_Vectorcall(func, args, nArgs | PY_VECTORCALL_ARGUMENTS_OFFSET, nullptr);
#endif
      return retVal;
    }
  }
}