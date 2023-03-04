#pragma once
// Include corecrt fixes this issue in pybind11 
// https://github.com/microsoft/onnxruntime/issues/9735
#include <corecrt.h>
#include "CPython.h"
#include <xloil/StringUtils.h>
#include <xloil/ExcelThread.h>
#include <xloil/Throw.h>
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <string>

/// Returns true if the object implements __iter__, compare with PyIter_Check which
/// tests if an object is an iterator. It seems useful, wonder why it's not in the API?
#define PyIterable_Check(obj) (obj)->ob_type->tp_iter != NULL


/// <summary>
/// Converts a PyObject to a str, then to a C++ string
/// </summary>
inline auto to_string(const PyObject* p)
{
  return (std::string)pybind11::str(pybind11::handle((PyObject*)p));
}
/// <summary>
/// Converts a PyObject to a str, then to a C++ wstring
/// </summary>
std::wstring to_wstring(const PyObject* p);

namespace pybind11
{
  /// <summary>
  /// A non-owning holder class used to bind references to static C++ objects
  /// </summary>
  template< typename T >
  class ReferenceHolder
  {
  public:
    explicit ReferenceHolder(T* ptr = nullptr) : ptr_(ptr) {}

    T* get() const { return ptr_; }
    T* operator-> () const { return ptr_; }

  private:
    T* ptr_;
  };

  inline auto to_string(const pybind11::object& p)
  {
    return to_string(p.ptr());
  }

  inline std::wstring to_wstring(const pybind11::object& p)
  {
    return to_wstring(p.ptr());
  }
}

PYBIND11_DECLARE_HOLDER_TYPE(T, pybind11::ReferenceHolder<T>, true);

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
    template<class TType = pybind11::object> inline TType PySteal(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_already_set();
      return pybind11::reinterpret_steal<TType>(obj);
    }
    template<class TType = pybind11::object> inline TType PyBorrow(PyObject* obj)
    {
      if (!obj)
        throw pybind11::error_already_set();
      return pybind11::reinterpret_borrow<TType>(obj);
    }

    /// <summary>
    /// Gets a proper reference to a weakref. Strangely, this functionality is missing
    /// in pybind11
    /// </summary>
    inline pybind11::object PyBorrow(const pybind11::weakref& wr)
    {
      return PyBorrow(PyWeakref_GetObject(wr.ptr()));
    }

    /// <summary>
    /// If PyErr_Occurred is true, returns the error message, else an empty string
    /// </summary>
    inline std::wstring pyErrIfOccurred(bool clear = true)
    {
      const auto result = PyErr_Occurred()
        ? utf8ToUtf16(pybind11::detail::error_fetch_and_normalize("").format_value_and_trace())
        : std::wstring();
      if (clear)
        PyErr_Clear();
      return result;
    }

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
    /// Returns a dangling reference
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

      void push_back(const pybind11::object& obj)
      {
        auto p = obj.ptr();
        Py_XINCREF(p);
        push_back(p);
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
        _size = TOffset;
      }
      PyObject* call(PyObject* func, PyObject* kwargs) noexcept
      {
        return fastCall(func, _store.data() + TOffset, nArgs(), kwargs);
      }

      pybind11::object call(const pybind11::object& func, const pybind11::object& kwargs)
      {
        return PySteal(fastCall(func.ptr(), _store.data() + TOffset, nArgs(), kwargs.ptr()));
      }
    };
  }
}