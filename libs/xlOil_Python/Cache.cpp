#include "xloil/ObjectCache.h"
#include "Cache.h"
#include "Main.h"
namespace py = pybind11;

namespace xloil {
  namespace Python {
    constexpr wchar_t thePyCacheUniquifier = L'\x6B23';
    static std::unique_ptr<ObjectCache<py::object, thePyCacheUniquifier>> thePythonObjCache;

    void createCache()
    {
      thePythonObjCache.reset(new ObjectCache<py::object, thePyCacheUniquifier>());
      static auto handler = Event_PyBye().bind([]() 
      {
        py::gil_scoped_acquire gil;
        thePythonObjCache.reset(); 
      });
    }

    ExcelObj addCache(py::object&& obj)
    {
      return thePythonObjCache->add(std::forward<py::object>(obj));
    }
    bool fetchCache(const wchar_t* cacheString, size_t length, py::object& obj)
    {
      return thePythonObjCache->fetch(cacheString, length, obj);
    }
} }