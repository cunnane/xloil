#pragma once
#include <memory>
namespace pybind11 { class handle; }
namespace xloil {
  class Application;
  namespace Python
  {
    class PyAddin;
  }
}

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// 'Hard' unloads a python module: clears its __dict__ and removes it
    /// from sys.modules. Release the module handle into the argument so
    /// there are no hanging references to the module object 
    /// </summary>
    bool unloadModule(const pybind11::handle& module);

    std::shared_ptr<const void>
      createWorkbookOpenHandler(const std::weak_ptr<PyAddin>& loadContext);
  }
}