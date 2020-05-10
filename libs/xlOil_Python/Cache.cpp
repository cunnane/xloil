#include "xloil/ObjectCache.h"
#include "BasicTypes.h"
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
    bool fetchCache(const std::wstring_view& str, py::object& obj)
    {
      return thePythonObjCache->fetch(str, obj);
    }

    namespace
    {
      py::object add_cache(py::object&& obj)
      {
        auto ref = addCache(std::forward<py::object>(obj));
        return PySteal<>(PyFromString()(ref));
      }
      static int theBinder = addBinder([](py::module& mod)
      {
        mod.def("to_cache", &add_cache);
      });
    }
} }