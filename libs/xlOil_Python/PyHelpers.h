#pragma once
// Horrible hack to allow our debug build to link with release python lib and so avoid building debug python
// Can remove this for Python >= 3.8
// https://stackoverflow.com/questions/17028576/using-python-3-3-in-c-python33-d-lib-not-found
#ifdef _DEBUG
#  define XLO_PY_HACK
#endif
#undef _DEBUG
#define HAVE_SNPRINTF
#include <Python.h>
#ifdef XLO_PY_HACK
#  define _DEBUG
#endif

#include "xloil/Utils.h"
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <string>

// Seems useful, wonder why it's not in the API?
#define PyIterable_Check(obj) \
    ((obj)->ob_type->tp_iter != NULL && \
     (obj)->ob_type->tp_iter != &_PyObject_NextNotImplemented)

namespace pybind11
{
  // Adds a logically missing wstr class to pybind11
  class wstr : public object {
  public:
    PYBIND11_OBJECT_CVT(wstr, object, detail::PyUnicode_Check_Permissive, raw_str)

    wstr(const wchar_t *c, size_t n)
      : object(PyUnicode_FromWideChar(c, (ssize_t)n), stolen_t{}) 
    {
      if (!m_ptr) pybind11_fail("Could not allocate string object!");
    }

    // 'explicit' is explicitly omitted from the following constructors to allow implicit 
    // conversion to py::str from C++ string-like objects
    wstr(const wchar_t *c = L"")
      : object(PyUnicode_FromWideChar(c, 0), stolen_t{})
    {
      if (!m_ptr) pybind11_fail("Could not allocate string object!");
    }

    wstr(const std::wstring &s) : wstr(s.data(), s.size()) { }

    // Not sure how to implement
    //explicit str(const bytes &b);

    /** \rst
    Return a string representation of the object. This is analogous to
    the ``str()`` function in Python.
    \endrst */
    explicit wstr(handle h) : object(raw_str(h.ptr()), stolen_t{}) { }

    operator std::wstring() const {
      if (!PyUnicode_Check(m_ptr))
        pybind11_fail("Unable to extract string contents!");
      ssize_t length;
      wchar_t* buffer = PyUnicode_AsWideCharString(ptr(), &length);
      return std::wstring(buffer, (size_t)length);
    }

    template <typename... Args>
    wstr format(Args &&...args) const {
      return attr("format")(std::forward<Args>(args)...);
    }

  private:
    /// Return string representation -- always returns a new reference, even if already a str
    static PyObject *raw_str(PyObject *op) {
      PyObject *str_value = PyObject_Str(op);
#if PY_MAJOR_VERSION < 3
      if (!str_value) throw error_already_set();
      PyObject *unicode = PyUnicode_FromEncodedObject(str_value, "utf-8", nullptr);
      Py_XDECREF(str_value); str_value = unicode;
#endif
      return str_value;
    }
  };
}

namespace xloil
{
  namespace Python
  {
    inline PyObject* PyCheck(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_already_set();
      return obj;
    }
    inline PyObject* PyCheck(int ret)
    {
      if (ret != 0)
        throw pybind11::error_already_set();
      return 0;
    }
    template<class TType = pybind11::object> inline TType PySteal(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_already_set();
      return pybind11::reinterpret_steal<TType>(obj);
    }
    template<class TType> inline TType PyBorrow(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_already_set();
      return pybind11::reinterpret_borrow<TType>(obj);
    }
    inline std::wstring pyErrIfOccurred()
    {
      return PyErr_Occurred() ? utf8ToUtf16(pybind11::detail::error_string().c_str()) : std::wstring();
    }
  }
}