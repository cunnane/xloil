#include "PyHelpers.h"

using std::wstring;

namespace xloil {
  namespace Python
  {
    bool sliceHelper1d(
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

    bool sliceHelper2d(
      const pybind11::tuple& loc,
      const size_t nRows, const size_t nCols,
      size_t& fromRow, size_t& fromCol,
      size_t& toRow, size_t& toCol)
    {
      if (loc.size() != 2)
        XLO_THROW("Expecting tuple of size 2");

      auto singleElement = sliceHelper1d(loc[0], nRows, fromRow, toRow);
      singleElement &= sliceHelper1d(loc[1], nCols, fromCol, toCol);
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