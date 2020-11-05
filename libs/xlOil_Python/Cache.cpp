#include <xloil/ObjectCache.h>
#include <xlOil/ExcelObjCache.h>
#include "PyCoreModule.h"
#include "BasicTypes.h"
#include "Cache.h"
#include "Main.h"
namespace py = pybind11;
using std::wstring;

namespace xloil {
  namespace Python {
    constexpr wchar_t thePyCacheUniquifier = L'\x6B23';

    namespace
    {
      struct PyCache;
      static PyCache* thePythonObjCache = nullptr;

      // Only a single instance of this class is created
      struct PyCache
      {
        PyCache()
          : _cache(false)
        {
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
          XLO_TRACE("Python object cache destroyed");
        }

        py::object add(const py::object& obj, const wchar_t* tag=nullptr)
        {
          // The cache expects callers to be of the form [.]xxx, so we add
          // a prefix if a custom tag is specified. Note the forward slash
          // cannot appear in a workbook name so this tag never collides with
          // the caller-based default
          const auto cacheKey = _cache.add(py::object(obj), tag 
              ? (wstring(L"[/Py/]") + tag).c_str()
              : nullptr);
          return PySteal(detail::PyFromString()(cacheKey.asPascalStr()));
        }
        py::object get(const std::wstring_view& str)
        {
          py::object obj;
          if (objectCacheCheckReference(str))
          {
            std::shared_ptr<const ExcelObj> xlObj;
            if (xloil::objectCacheFetch(str, xlObj))
              return PySteal(PyFromAny()(*xlObj));
          }
          else
            _cache.fetch(str, obj);
          return obj;
        }
        bool remove(const std::wstring& cacheRef)
        {
          return _cache.remove(cacheRef);
        }
        bool contains(const std::wstring_view& str)
        {
          py::object obj;
          return _cache.fetch(str, obj);
        }

        py::list keys() const
        {
          py::list out;
          for (auto&[wbName, wbCache] : _cache)
            for (auto&[address, cellCache] : *wbCache)
              for (auto i = 0; i < cellCache->objects().size(); ++i)
                out.append(py::wstr(_cache.writeKey(wbName, address, i)));
          return out;
        }

        ObjectCache<py::object, thePyCacheUniquifier> _cache;
        std::shared_ptr<const void> _workbookCloseHandler;
      };
    }
    ExcelObj pyCacheAdd(const py::object& obj, const wchar_t* caller)
    {
      return thePythonObjCache->_cache.add(py::object(obj), caller);
    }
    bool pyCacheGet(const std::wstring_view& str, py::object& obj)
    {
      return thePythonObjCache->_cache.fetch(str, obj);
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyCache>(mod, "ObjectCache")
          .def("add", &PyCache::add, py::arg("obj"), py::arg("tag")=nullptr)
          .def("remove", &PyCache::remove, py::arg("ref"))
          .def("get", &PyCache::get, py::arg("ref"))
          .def("contains", &PyCache::contains, py::arg("ref"))
          .def("keys", &PyCache::keys)
          .def("__contains__", &PyCache::contains)
          .def("__getitem__", &PyCache::get)
          .def("__call__", &PyCache::add, py::arg("obj"), py::arg("tag") = nullptr);
        mod.add_object("cache", py::cast(new PyCache()));
      });
    }
} }