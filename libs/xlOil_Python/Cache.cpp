#include <xloil/ObjectCache.h>
#include <xlOil/ExcelObjCache.h>
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
        {
          thePythonObjCache = this;
        }

        // Just to prevent any potential errors!
        PyCache(const PyCache& that) = delete;

        ~PyCache()
        {
          XLO_TRACE("Python object cache destroyed");
        }

        py::object add(const py::object& obj)
        {
          return PySteal(PyFromString()(_cache.add(py::object(obj))));
        }
        py::object get(const std::wstring_view& str)
        {
          py::object obj;
          if (objectCacheCheckReference(str))
          {
            std::shared_ptr<const ExcelObj> xlObj;
            if (xloil::objectCacheFetch(str, xlObj))
              return PySteal(PyFromAny<>()(*xlObj));
          }
          else
            _cache.fetch(str, obj);
          return obj;
        }
        bool contains(const std::wstring_view& str)
        {
          py::object obj;
          return _cache.fetch(str, obj);
        }

        ObjectCache<py::object, thePyCacheUniquifier> _cache;
      };
    }
    ExcelObj pyCacheAdd(py::object&& obj)
    {
      return thePythonObjCache->_cache.add(std::forward<py::object>(obj));
    }
    bool pyCacheGet(const std::wstring_view& str, py::object& obj)
    {
      //return false;
      return thePythonObjCache->_cache.fetch(str, obj);
    }

    namespace
    {
      static int theBinder = addBinder([](py::module& mod)
      {
        py::class_<PyCache>(mod, "ObjectCache")
          .def("add", &PyCache::add)
          .def("get", &PyCache::get)
          .def("contains", &PyCache::contains)
          .def("__contains__", &PyCache::contains)
          .def("__getitem__", &PyCache::get)
          .def("__call__", &PyCache::add);
        mod.add_object("cache", py::cast(new PyCache()));
      });
    }
} }