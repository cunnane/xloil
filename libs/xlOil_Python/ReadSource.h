#pragma once
namespace pybind11 { class object; class module; }

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// Calls scan_module in the xloil.py file on the specified module.
    /// This function looks for appropriately decorated xlOil functions
    /// and classes to register. It can be called repeatedly on the same 
    /// module.
    /// </summary>
    void scanModule(const pybind11::object& mod);

    pybind11::object loadModuleFromFile(const wchar_t* filepath, const wchar_t* linkedWorkbook = nullptr);

    /// <summary>
    /// 'Hard' unloads a python module: clears its __dict__ and removes it
    /// from sys.modules
    /// </summary>
    bool unloadModule(const pybind11::module& module);

    void createWorkbookOpenHandler(const wchar_t* starredPattern);
  }
}