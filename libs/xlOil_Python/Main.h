#pragma once
#include <memory>

namespace xloil
{
  class AddinContext; class FileSource;

  namespace Python
  {
    class EventLoop;

    /// <summary>
    /// Hold a python addin context. Each XLL which uses xlOil_Python has a 
    /// separate context to keep track of the functions it registers. It also
    /// has separate thread and event loop on which all importing is done
    /// </summary>
    struct PyAddin
    {
      PyAddin(AddinContext&);
      AddinContext& context;
      std::unique_ptr<EventLoop> thread;
    };

    PyAddin& findAddin(const wchar_t* xllPath);
    PyAddin& theCurrentAddin();

    /// <summary>
    /// The core context corresponds to xlOil.dll - it always exists and is
    /// used for loading any modules specified in the core settings and addin 
    /// non-specific stuff such as workbook modules and jupyter functions. 
    /// </summary>
    /// <returns></returns>
    PyAddin& theCoreAddin();

    /// <summary>
    /// Similar to the function in FileSource, but retrieve the PyAddin
    /// instead of the AddinContext
    /// </summary>
    /// <param name="sourcePath"></param>
    /// <returns></returns>
    std::pair<std::shared_ptr<FileSource>, PyAddin*> 
      findSource(const wchar_t* sourcePath);

   }
}