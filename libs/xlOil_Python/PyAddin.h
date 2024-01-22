#pragma once
#include <pybind11/pybind11.h>
#include <memory>
#include <string>
#include <map>

namespace xloil
{
  class AddinContext; class FuncSource;
  namespace Python { class EventLoop; }
}

namespace xloil
{
  namespace Python
  {
    /// <summary>
    /// Holds a python addin context. Each XLL which uses xlOil_Python has a 
    /// separate context to keep track of the functions it registers. It also
    /// has separate thread and event loop on which all importing is done
    /// </summary>
    class PyAddin : public std::enable_shared_from_this<PyAddin>
    {
    public:
      PyAddin(
        AddinContext&, 
        bool separateThread,
        const std::string_view& comLib, 
        const std::wstring_view& wbPattern = std::wstring_view());

      /// <summary>
      /// "Guts" the addin object in the case that a detach message is received
      /// from xlOil core. This ensures any remaining references fail gracefully
      /// </summary>
      void unload();

      std::shared_ptr<EventLoop> thread;
     
      /// <summary>
      /// Gets the addin pathname
      /// </summary>
      const std::wstring& pathName() const;

      /// <summary>
      /// Given a workbook path, returns the expected location of its local 
      /// module (i.e. py file), based on the pattern specified in the ini file.
      /// </summary>
      std::wstring getLocalModulePath(const wchar_t* workbookPath) const;

      /// <summary>
      /// Imports / reloads the specified modules and scans them for functions
      /// to register. The argument is passed to `xloil.importer._import_and_scan`
      /// so a module, string or enumerable of the these can be given. 
      /// </summary>
      void importModule(const pybind11::object& module);

      /// <summary>
      /// Imports the specified py file without registering it as module in 
      /// `sys.modules`, then scans for functions to register.  Optionally
      /// specifies a linked workbook which is passed back when functions are
      /// registered
      /// </summary>
      void importFile(const wchar_t* filePath, const wchar_t* linkedWorkbook);

      std::shared_ptr<FuncSource> findSource(const wchar_t* sourcePath) const;
      
      bool loadLocalModules() const { return !_workbookPattern.empty(); }

      /// <summary>
      /// Returns the python com support library specified in the addin's 
      /// settings, e.g. 'win32com'
      /// </summary>
      const std::string& comBinder() const { return _comBinder; }

      AddinContext& context();
      const AddinContext& context() const;

    private:
      std::wstring _workbookPattern;
      AddinContext* _context;
      std::string _comBinder;
      pybind11::object self() const;
    };

    /// <summary>
    /// Only called from Main.cpp on plugin startup
    /// </summary>
    std::map<std::wstring, std::shared_ptr<PyAddin>>& getAddins();

    /// <summary>
    /// Returns the PyAddin object corresponding to the given XLL,
    /// or throws if none found.
    /// </summary>
    PyAddin& findAddin(const wchar_t* xllPath);
    
    /// <summary>
    /// Gets the event loop associated with the current thread or throws
    /// </summary>
    /// <returns></returns>
    std::shared_ptr<EventLoop> getEventLoop();

    /// <summary>
    /// The core context corresponds to xlOil.dll - it always exists and is
    /// used for loading any modules specified in the core settings and addin 
    /// non-specific stuff such as workbook modules and jupyter functions. 
    /// </summary>
    /// <returns></returns>
    const std::shared_ptr<PyAddin>& theCoreAddin();
  }
}