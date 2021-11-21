#pragma once
namespace pybind11 { class object; class handle; }

namespace xloil
{
  namespace Python
  {
    struct PyAddin;

    /// <summary>
    /// Calls scan_module in the xloil.py file on the specified module.
    /// This function looks for appropriately decorated xlOil functions
    /// and classes to register. It can be called repeatedly on the same 
    /// module.
    /// </summary>
    void scanModule(const pybind11::object& mod);

    /// <summary>
    /// 'Hard' unloads a python module: clears its __dict__ and removes it
    /// from sys.modules. Release the module handle into the argument so
    /// there are no hanging references to the module object 
    /// </summary>
    bool unloadModule(const pybind11::handle& module);

    void createWorkbookOpenHandler(const wchar_t* starredPattern, PyAddin& loadContext);
  }
}