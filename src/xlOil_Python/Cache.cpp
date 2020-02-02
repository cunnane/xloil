#include "xloil/ObjectCache.h"
#include "Cache.h"
#include "Main.h"

namespace xloil {
  namespace Python {
    static std::unique_ptr<ObjectCache<PyObject*>> thePythonObjCache;

    void createCache()
    {
      thePythonObjCache.reset(new ObjectCache<PyObject*>(L'\x6B23'));
      static auto handler = Event_PyBye().bind([]() {thePythonObjCache.reset(); });
    }

    ExcelObj addCache(PyObject* obj)
    {
      return thePythonObjCache->add(obj);
    }
    bool fetchCache(const wchar_t* cacheString, size_t length, PyObject*& obj)
    {
      return thePythonObjCache->fetch(cacheString, length, obj);
    }
} }