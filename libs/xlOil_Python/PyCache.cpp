#include <xlOil/ExcelObjCache.h>
#include <xlOil/ObjectCache.h>
#include "PyCore.h"
#include "TypeConversion/BasicTypes.h"
#include "PyCache.h"

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
      /// <summary>
      /// This odd singleton is constructed and owned by the core module which ensures
      /// deleted when the core module is garbage collected and the interpreter is 
      /// still active and GIL is held.
      /// </summary>
      class PyCache
      {
        using cache_type = ObjectCache<py::object, CacheUniquifier<py::object>>;

        PyCache()
          : _cache(cache_type::create(false))
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
          XLO_DEBUG("Python object cache destroyed");
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

        py::object add(py::object& obj, const wstring& tag, const wstring& key)
        {
          const auto cacheKey = key.empty()
            ? _cache->add(std::move(obj), CallerInfo(), tag)
            : _cache->add(std::move(obj), key);
          return PySteal(detail::PyFromString()(cacheKey.cast<PStringRef>()));
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
        bool remove(const std::wstring& cacheRef)
        {
          return _cache->erase(cacheRef);
        }
        bool contains(const std::wstring_view& str)
        {
          return _cache->fetch(str);
        }

        py::list keys() const
        {
          py::list out;
          for (auto& [key, cellCache] : *_cache)
            for (uint16_t i = 0u; i < cellCache.count(); ++i)
              out.append(py::wstr(_cache->writeKey(key, i)));
          return out;
        }

        shared_ptr<cache_type> _cache;
        shared_ptr<const void> _workbookCloseHandler;
      };

      PyCache* PyCache::_theInstance = nullptr;
    }

    ExcelObj pyCacheAdd(const py::object& obj, const wchar_t* caller)
    {
      // Decorate the cache ref with the python object name to 
      // help users keep track
      auto name = utf8ToUtf16(obj.ptr()->ob_type->tp_name);
      return PyCache::instance()._cache->add(
        py::object(obj),
        caller ? CallerInfo(ExcelObj(caller)) : CallerInfo(),
        name);
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
        py::class_<PyCache>(mod, "ObjectCache", R"(
            Provides a way to manipulate xlOil's Python object cache

            Examples
            --------

            ::
        
                @xlo.func
                def myfunc(x):
                    return xlo.cache(MyObject(x)) # <-equivalent to cache.add(...)

                @xlo.func
                def myfunc2(array: xlo.Array(str), i):
                    return xlo.cache[array[i]]   # <-equivalent to cache.get(...)

          )")
          .def("add", 
            &PyCache::add, 
            R"(
              Adds an object to the cache and returns a reference string.

              xlOil automatically adds objects returned from worksheet 
              functions to the cache if they cannot be converted by any 
              registered converter.  So this function is useful to:
              
                 1) force a convertible object, such as an iterable, into the
                    cache
                 2) return a list of cached objects
                 3) create cached objects from outside of worksheet fnctions
                    e.g. in commands / subroutines

              xlOil uses the caller infomation provided by Excel to construct
              the cache string and manage the cache object lifecycle. When
              invoked from a worksheet function, this caller info contains 
              the cell reference. xlOil deletes cache objects linked to the 
              cell reference from previous calculation cycles.

              When invoked from a source other than a worksheet function (there
              are several possibilies, see the help for `xlfCaller`), xlOil
              again generates a reference string based on the caller info. 
              However, this may not be unique.  In addition, objects with the 
              same caller string will replace those created during a previous 
              calculation cycle. For example, creating cache objects from a button
              clicked repeatedly will behave differently if Excel recalculates 
              in between the clicks. To override this behaviour, the exact cache
              `key` can be specified.  For example, use Python's `id` function or
              the cell address being written to if a command is writing a cache
              string to the sheet.  When `key` is specified the user is responsible
              for managing the lifecycle of their cache objects.
 

              Parameters
              ----------

              obj:
                The object to cache.  Required.

              tag: str
                An optional string to append to the cache ref to make it more 
                'friendly'. When returning python objects from functions, 
                xlOil uses the object's type name as a tag

              key: str
                If specified, use the exact cache key (after prepending by
                cache uniquifier). The user is responsible for ensuring 
                uniqueness of the cache key.
            )",
            py::arg("obj"), py::arg("tag") = "", py::arg("key")="")
          .def("remove", &PyCache::remove, py::arg("ref"),
            R"(
              xlOil manages the lifecycle for most cache objects, so this  
              function should only be called when `add` was invoked with a
              specified key - in this case the user owns the lifecycle 
              management. 
            )")
          .def("get", 
            &PyCache::get, 
            R"(
              Fetches an object from the cache given a reference string.
              Returns `default` if not found
            )",
            py::arg("ref"), py::arg("default") = py::none())
          .def("contains", 
            &PyCache::contains, 
            "Returns True if the given reference string links to a valid object",
            py::arg("ref"))
          .def("keys", 
            &PyCache::keys,
            "Returns all cache keys as a list of strings")
          .def("__contains__", &PyCache::contains)
          .def("__getitem__", &PyCache::getitem)
          .def("__call__", 
            &PyCache::add, 
            "Calls `add` method with provided arguments",
            py::arg("obj"), py::arg("tag")="", py::arg("key")="");

        mod.add_object("cache", py::cast(PyCache::construct()));
      });
    }
  }
}