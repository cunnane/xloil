#pragma once
#include "CPython.h"
#include <xloil/StringUtils.h>
#include <xloil/ExcelThread.h>
#include <xloil/Throw.h>
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
    PYBIND11_OBJECT_CVT(wstr, object, PYBIND11_STR_CHECK_FUN, raw_str)

    wstr(const wchar_t* c, size_t n)
      : object(PyUnicode_FromWideChar(c, (ssize_t)n), stolen_t{})
    {
      if (!m_ptr)
        pybind11_fail("Could not allocate string object!");
    }

    // 'explicit' is omitted from the following constructors to allow implicit 
    // conversion to py::str from C++ string-like objects
    wstr(const wchar_t* c = L"")
      : object(PyUnicode_FromWideChar(c, -1), stolen_t{})
    {
      if (!m_ptr)
        pybind11_fail("Could not allocate string object!");
    }

    wstr(const std::wstring_view& s) : wstr(s.data(), s.size()) { }

    // Not sure how to implement
    //explicit str(const bytes &b);

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
    static PyObject* raw_str(PyObject* op) {
      PyObject* str_value = PyObject_Str(op);
      return str_value;
    }
  };


  /// <summary>
  /// Provides a replacement for pybind's detail::error_string which handles
  /// the auxillary context and cause expceptions.
  /// </summary>
  /// <returns></returns>
  std::string error_full_traceback();

  class error_traceback_set : public error_already_set
  {
  public:
    // Note: When pybind is upgraded, we need to add a ctor to 
    // error_already_set which takes a string msg
    error_traceback_set()
      : error_already_set(error_full_traceback())
    {}
  };
}

namespace xloil
{
  namespace Python
  {
    inline PyObject* PyCheck(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_traceback_set();
      return obj;
    }
    template<class TType = pybind11::object> inline TType PySteal(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_traceback_set();
      return pybind11::reinterpret_steal<TType>(obj);
    }
    template<class TType = pybind11::object> inline TType PyBorrow(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_traceback_set();
      return pybind11::reinterpret_borrow<TType>(obj);
    }

    /// <summary>
    /// If PyErr_Occurred is true, returns the error message, else an empty string
    /// </summary>
    inline std::wstring pyErrIfOccurred(bool clear = true)
    {
      const auto result = PyErr_Occurred()
        ? utf8ToUtf16(pybind11::error_full_traceback())
        : std::wstring();
      if (clear)
        PyErr_Clear();
      return result;
    }

    /// <summary>
    /// Converts a PyObject to a str, then to a C++ string
    /// </summary>
    inline auto pyToStr(const PyObject* p)
    {
      // Is morally const: py::handle doesn't change refcount
      return (std::string)pybind11::str(pybind11::handle((PyObject*)p));
    }

    /// <summary>
    /// Converts a PyObject to a str, then to a C++ wstring
    /// </summary>
    std::wstring pyToWStr(const PyObject* p);

    inline std::wstring
      pyToWStr(const pybind11::object& p) { return pyToWStr(p.ptr()); }

    /// <summary>
    /// Reads an argument to __getitem__ i.e. [] using the following rules
    ///     None => entire array
    ///     Slice [a:b] => compute indices using python rules
    ///     int => single value (0-based)
    /// Modifies the <param ref="from"/> and <param ref="to"/> arguments
    /// to indicate the extent of the sliced array. Only handles slices with
    /// stride = 1.
    /// </summary>
    /// <param name="index"></param>
    /// <param name="size">The size of the object being indexed</param>
    /// <param name="from"></param>
    /// <param name="to"></param>
    /// <returns>Returns true if only a single element is accessed</returns>
    bool getItemIndexReader1d(
      const pybind11::object& index,
      const size_t size, size_t& from, size_t& to);

    /// <summary>
    /// Take a 2-tuple of indeices and applies <see ref="getItemIndexReader1d"/> in 
    /// each dimension
    /// </summary>
    /// <param name="loc"></param>
    /// <param name="nRows">The first dimension of the object being indexed</param>
    /// <param name="nCols">The second dimension of the object being indexed</param>
    /// <param name="fromRow"></param>
    /// <param name="fromCol"></param>
    /// <param name="toRow"></param>
    /// <param name="toCol"></param>
    /// <returns>Returns true if only a single element is accessed/returns>
    bool getItemIndexReader2d(
      const pybind11::tuple& loc,
      const size_t nRows, const size_t nCols,
      size_t& fromRow, size_t& fromCol,
      size_t& toRow, size_t& toCol);

    /// <summary>
    /// Holds a py::object and ensures the GIL is held when the holder is destroyed
    /// and the underlying py::object is decref'd 
    /// </summary>
    class PyObjectHolder : public pybind11::detail::object_api<PyObjectHolder>
    {
      pybind11::object _obj;
    public:
      PyObjectHolder(const pybind11::object& obj)
        : _obj(obj)
      {}
      ~PyObjectHolder()
      {
        if (!_obj)
          return;
        pybind11::gil_scoped_acquire getGil;
        _obj = std::move(pybind11::object());
      }
      operator pybind11::object() const { return _obj; }

      /// Return the underlying ``PyObject *`` pointer
      PyObject* ptr() const { return _obj.ptr(); }
      PyObject*& ptr() { return _obj.ptr(); }
    };

    /// <summary>
    /// Wraps a class member function to ensure the GIL is released before it
    /// is called.  Used for pybind: e.g. mod.def("bar", wrapNoGil(&Foo::bar))
    /// </summary>
    template<class Return, class Class, class... Args>
    constexpr auto wrapNoGil(Return(Class::* f)(Args...) const)
    {
      return [f](Class* self, Args... args)
      {
        py::gil_scoped_release release;
        return (self->*f)(args...);
      };
    }

    template<class Return, class Class, class... Args>
    constexpr auto wrapNoGil(Return(Class::* f)(Args...))
    {
      return [f](Class* self, Args... args)
      {
        py::gil_scoped_release release;
        return (self->*f)(args...);
      };
    }

    template<class F>
    constexpr auto wrapNoGil(F&& f)
    {
      py::gil_scoped_release release;
      return f();
    }

    /// <summary>
    /// Wraps a class member function to ensure it is executed on Excel's main
    /// thread (with no GIL) Used for pybind: e.g. mod.def("bar", MainThreadWrap(&Foo::bar))
    /// </summary>
    template<class Return, class Class, class... Args>
    constexpr auto MainThreadWrap(Return(Class::* f)(Args...) const)
    {
      return [f](Class* self, Args... args)
      {
        auto fut = runExcelThread([=]()
        {
          return (self->*f)(args...);
        });
        py::gil_scoped_release release;
        return fut.get();
      };
    }

    template<class Return, class Class, class... Args>
    constexpr auto MainThreadWrap(Return(Class::* f)(Args...))
    {
      return [f](Class* self, Args... args)
      {
        auto fut = runExcelThread([=]()
        {
          return (self->*f)(args...);
        });
        py::gil_scoped_release release;
        return fut.get();
      };
    }

    template<class F, class T, class Return, class... Args>
    constexpr auto MainThreadWrap(
      F&& f,
      Return(T::*)(Args...) const)
    {
      return [f](Args... args)
      {
        auto fut = runExcelThread([=]()
        {
          return f(args...);
        });
        py::gil_scoped_release release;
        return fut.get();
      };
    }

    template<class F>
    constexpr auto MainThreadWrap(F&& f)
    {
      return MainThreadWrap(f, (decltype(&F::operator())) nullptr);
    }

    /// <summary>

    /// </summary>
    PyObject* fastCall(
      PyObject* func, PyObject* const* args, size_t nArgs, PyObject* kwargs) noexcept;

    /// <summary>
    /// Manages an array of args suitable for a Python FastCall, this includes the 
    /// leading offset to allow easy fiddling of the 'self' parameter for onward calls.
    /// Python can optimise onward calls to PyObject_Vectorcall if we leave a free 
    /// entry at the start of the arg array For Py 3.7 and earlier, vector call is not
    /// available.
    /// 
    /// The array is held on the stack, so a maximum size must be specified. 
    /// </summary>
    template<
      size_t TSize = 255, // = XL_MAX_UDF_ARGS
#if PY_VERSION_HEX < 0x03080000
      size_t TOffset = 1u
#else
      size_t TOffset = 0u
#endif
    >
    class PyCallArgs
    {
      // Use array<PyObject*> as an array<py::object> would result in TSize dtor calls
      std::array<PyObject*, TSize + TOffset>  _store;
      size_t _size = TOffset;

    public:
      ~PyCallArgs()
      {
        clear();
      }

      /// <summary>
      /// Steals a ref
      /// </summary>
      void push_back(PyObject* p)
      {
        assert(_size <= TSize);
        _store[_size++] = p;
      }

      constexpr auto begin() const
      {
        return _store.begin();
      }

      auto end() const
      {
        return begin() + _size;
      }

      size_t nArgs() const { return _size - TOffset; }

      constexpr size_t capacity() const { return TSize; }

      void clear()
      {
        const auto last = end();
        for (auto p = _store.begin() + TOffset; p != last; ++p)
          Py_DECREF(*p);
        _size = 0;
      }

      PyObject* call(PyObject* func, PyObject* kwargs) noexcept
      {
        return fastCall(func, _store.data() + TOffset, nArgs(), kwargs);
      }
    };
  }
}