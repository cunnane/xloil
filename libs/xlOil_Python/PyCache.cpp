#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCache.h"
#include <xlOil/ExcelObjCache.h>
#include <xlOil/ObjectCache.h>

namespace py = pybind11;
using std::wstring;
using std::shared_ptr;

namespace xloil
{
  template<>
  struct CacheUniquifier<py::object>
  {
    static constexpr wchar_t value = L'\x6B23';
  };
  using pyCacheUnquifier = CacheUniquifier<py::object>;

  namespace Python {

    namespace
    {
      struct PyObjToPtr
      {
        intptr_t operator()(const py::object& obj) const
        {
          return (intptr_t)obj.ptr();
        }
      };

      /// <summary>
      /// This odd singleton is constructed and owned by the core module which ensures
      /// deleted when the core module is garbage collected and the interpreter is 
      /// still active and GIL is held.
      /// </summary>
      class PyCache
      {
        using cache_type = ObjectCache<PyObjectHolder, CacheUniquifier<py::object>, PyObjToPtr>;

        PyCache()
          : _cache(cache_type::create())
        {
          _workbookCloseHandler = std::static_pointer_cast<const void>(
            xloil::Event::WorkbookAfterClose().bind(
              [this](auto wbName)
          {
            py::gil_scoped_acquire getGil;
            _cache->onWorkbookClose(wbName);
          }));
        }

        // Just to prevent any potential errors!
        PyCache(const PyCache& that) = delete;

        static PyCache* _theInstance;

      public:

        ~PyCache()
        {
          _theInstance = nullptr;
          XLO_TRACE("Python object cache destroyed");
        }

        static PyCache* construct()
        {
          _theInstance = new PyCache();
          return _theInstance;
        }

        static PyCache& instance()
        {
          assert(_theInstance);
          return *_theInstance;
        }

        py::object add(py::object& obj, const wstring& tag, const wstring& caller)
        {
          auto callerInfo = caller.empty() ? CallerInfo() : CallerInfo(ExcelObj(caller));
          const auto cacheKey = _cache->add(std::move(obj), std::move(callerInfo), tag);
          return PySteal(detail::PyFromString()((PStringView<>)cacheKey));
        }
        py::object getitem(const std::wstring_view& str)
        {
          auto result = get(str);
          if (result.is_none())
            throw pybind11::key_error(utf16ToUtf8(str));
          return result;
        }
        py::object get(const std::wstring_view& str, const py::object& default = py::none())
        {
          const ExcelObj* xlObj = getCached<ExcelObj>(str);
          if (xlObj)
            return PySteal(PyFromAny()(*xlObj));

          auto* obj = _cache->fetch(str);
          return obj ? *obj : default;
        }
        bool contains(const std::wstring_view& str)
        {
          return _cache->fetch(str);
        }

        py::list keys() const
        {
          py::list out;
          for (auto& [key, val] : *_cache)
            out.append(py::wstr(_cache->keyToStr(key).view()));
          return out;
        }

        shared_ptr<cache_type> _cache;
        shared_ptr<const void> _workbookCloseHandler;
      };

      PyCache* PyCache::_theInstance = nullptr;
    }

    ExcelObj pyCacheAdd(const py::object& obj, CallerInfo&& caller)
    {
      // Decorate the cache ref with the python object name to 
      // help users keep track of their objects
      auto name = utf8ToUtf16(obj.ptr()->ob_type->tp_name) + L' ';
      return ExcelObj(PyCache::instance()._cache->add(
        py::object(obj),
        std::move(caller),
        name));
    }

    bool pyCacheGet(const std::wstring_view& str, py::object& obj)
    {
      auto& cache = *PyCache::instance()._cache;
      if (!cache.valid(str))
        return false;

      const auto* p = cache.fetch(str);
      if (!p)
        return false;

      obj = *p;
      return true;
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyCache>(mod, "ObjectCache")
          .def("add", &PyCache::add, py::arg("obj"), py::arg("tag") = "", py::arg("key")="")
          .def("remove", &PyCache::remove, py::arg("ref"))
          .def("get", &PyCache::get, py::arg("ref"), py::arg("default"))
          .def("contains", &PyCache::contains, py::arg("ref"))
          .def("keys", &PyCache::keys)
          .def("__contains__", &PyCache::contains)
          .def("__getitem__", &PyCache::getitem)
          .def("__call__", 
            &PyCache::add, 
            "Calls `add` method with provided arguments",
            py::arg("obj"), py::arg("tag")="", py::arg("key")="");

        mod.add_object("cache", py::cast(PyCache::construct(), py::return_value_policy::take_ownership));
      });
    }
  }
}