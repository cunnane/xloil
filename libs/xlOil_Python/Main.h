#pragma once
#include "PyAddin.h"
#include <memory>
#include <string>

namespace xloil
{
  class FuncSource;

  namespace Python
  {
    /// <summary>
    /// The core context corresponds to xlOil.dll - it always exists and is
    /// used for loading any modules specified in the core settings and addin 
    /// non-specific stuff such as workbook modules and jupyter functions. 
    /// </summary>
    /// <returns></returns>
    const std::shared_ptr<PyAddin>& theCoreAddin();

    /// <summary>
    /// Similar to the function in FileSource, but retrieve the PyAddin
    /// instead of the AddinContext
    /// </summary>
    /// <param name="sourcePath"></param>
    /// <returns></returns>
    std::pair<std::shared_ptr<FuncSource>, PyAddin*> 
      findSource(const wchar_t* sourcePath);
   }
}