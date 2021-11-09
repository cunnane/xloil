#include <xlOil/ExcelObjCache.h>
#include <xlOil/ObjectCache.h>
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCache.h"
//#include "Main.h"

namespace py = pybind11;
using std::wstring;

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
      struct PyCache;
      // Non-owning pointer
      static PyCache* thePythonObjCache = nullptr;

      // Only a single instance of this class is created
      struct PyCache
      {
        PyCache()
          : _cache(false)
        {
          static_assert(CACHE_KEY_MAX_LEN == decltype(PyCache::_cache)::KEY_MAX_LEN);

          thePythonObjCache = this;
          _workbookCloseHandler = std::static_pointer_cast<const void>(
            xloil::Event::WorkbookAfterClose().bind(
              [this](auto wbName)
              { 
                py::gil_scoped_acquire getGil;
                _cache.onWorkbookClose(wbName); 
              }));
        }

        // Just to prevent any potential errors!
        PyCache(const PyCache& that) = delete;

        ~PyCache()
        {
          thePythonObjCache = nullptr;
          XLO_TRACE("Python object cache destroyed");
        }

        py::object add(py::object obj, const wchar_t* tag=nullptr)
        {
          // The cache expects callers to be of the form [.]xxx, so we add
          // a prefix if a custom tag is specified. Note the forward slash
          // cannot appear in a workbook name so this tag never collides with
          // the caller-based default
          const auto cacheKey = _cache.add(std::move(obj), tag 
              ? CallerLite(ExcelObj(tag))
              : CallerLite());
          return PySteal(detail::PyFromString()(cacheKey.asPString()));
        }
        py::object getitem(const std::wstring_view& str)
        {
          auto result = get(str);
          if (result.is_none())
            throw pybind11::key_error(utf16ToUtf8(str));
          return result;
        }
        py::object get(const std::wstring_view& str, const py::object& default=py::none())
        {
          const ExcelObj* xlObj = getCached<ExcelObj>(str);
          if (xlObj)
            return PySteal(PyFromAny()(*xlObj));

          auto* obj = _cache.fetch(str);
          return obj ? *obj : default;
        }
        bool remove(const std::wstring& cacheRef)
        {
          return _cache.erase(cacheRef);
        }
        bool contains(const std::wstring_view& str)
        {
          return _cache.fetch(str);
        }

        py::list keys() const
        {
          py::list out;
          for (auto&[key, cellCache] : _cache)
            for (auto i = 0u; i < cellCache.count(); ++i)
              out.append(py::wstr(_cache.writeKey(key, i)));
          return out;
        }

        ObjectCache<py::object, CacheUniquifier<py::object>> _cache;
        std::shared_ptr<const void> _workbookCloseHandler;
      };
    }
    ExcelObj pyCacheAdd(const py::object& obj, const wchar_t* caller)
    {
      if (!thePythonObjCache)
        XLO_THROW("Fatal: Python object cache not available");
      auto name = utf8ToUtf16(obj.ptr()->ob_type->tp_name);
      return thePythonObjCache->_cache.add(
        py::object(obj),
        caller ? CallerLite(ExcelObj(caller)) : CallerLite(),
        name.c_str(),
        name.size());
    }
    bool pyCacheGet(const std::wstring_view& str, py::object& obj)
    {
      if (!thePythonObjCache)
        XLO_THROW("Fatal: Python object cache not available");
      const auto* p = thePythonObjCache->_cache.fetch(str);
      if (p)
        obj = *p;
      return p;
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyCache>(mod, "ObjectCache")
          .def("add", &PyCache::add, py::arg("obj"), py::arg("tag")=nullptr)
          .def("remove", &PyCache::remove, py::arg("ref"))
          .def("get", &PyCache::get, py::arg("ref"), py::arg("default"))
          .def("contains", &PyCache::contains, py::arg("ref"))
          .def("keys", &PyCache::keys)
          .def("__contains__", &PyCache::contains)
          .def("__getitem__", &PyCache::getitem)
          .def("__call__", &PyCache::add, py::arg("obj"), py::arg("tag") = nullptr);
        mod.add_object("cache", py::cast(new PyCache()));
      });
    }
} }