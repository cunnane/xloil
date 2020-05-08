#include "Interface.h"
#include <xlOil/Register/FuncRegistry.h>
#include <xlOil/Register/LocalFunctions.h>
#include "ObjectCache.h"
#include <xlOilHelpers/Settings.h>
#include "EntryPoint.h"
#include "Log.h"
#include <xlOil/Loaders/AddinLoader.h>
#include <ComInterface/Connect.h>
#include <toml11/toml.hpp>

using std::make_pair;
using std::wstring;
using std::make_shared;
using std::shared_ptr;

namespace xloil
{


  const wchar_t* Core::theCorePath()
  {
    return xloil::theCorePath();
  }
  const wchar_t* Core::theCoreName()
  {
    return xloil::theCoreName();
  }

  int Core::theExcelVersion()
  {
    return coreExcelVersion();
  }

  Excel::_Application& Core::theExcelApp()
  {
    return excelApp();
  }

  bool Core::inFunctionWizard()
  {
    return xloil::inFunctionWizard();
  }

  void Core::throwInFunctionWizard()
  {
    if (xloil::inFunctionWizard())
      XLO_THROW("#WIZARD!");
  }

  AddinContext::AddinContext(
    const wchar_t* pathName, std::shared_ptr<const toml::value> settings)
    : _settings(settings)
    , _pathName(pathName)
  {
  }

  AddinContext::~AddinContext()
  {
  }

  FileSource::FileSource(const wchar_t* sourceName, bool watchSource)
    : _source(sourceName)
  {
    //_functionPrefix = toml::find_or<std::string>(*_settings, "FunctionPrefix", "");
    // TODO: watch source
    //Event_DirectoryChange()
  }

  FileSource::~FileSource()
  {
    XLO_DEBUG(L"Deregistering functions in source '{0}'", _source);
    forgetLocalFunctions(_workbookName.c_str());
    for (auto& f : _functions)
      xloil::deregisterFunc(f.second);
    _functions.clear();
  }

  bool FileSource::registerFuncs(
    std::vector<std::shared_ptr<const FuncSpec> >& funcSpecs)
  {
    decltype(_functions) newFuncs;

    for (auto& f : funcSpecs)
    {
      // If registration succeeds, just add the function to the new map
      auto ptr = registerFunc(f);
      if (ptr)
      {
        _functions.erase(f->name());
        newFuncs.emplace(f->name(), ptr);
        f.reset();
      }
    }

    // Remove all the null FuncSpec ptrs
    funcSpecs.erase(
      std::remove_if(funcSpecs.begin(), funcSpecs.end(), [](auto& f) { return !f; }),
      funcSpecs.end());

    // Any functions remaining in the old map must have been removed from the module
    // so we can deregister them, but if that fails we have to keep them or they
    // will be orphaned
    for (auto& f : _functions)
      if (!xloil::deregisterFunc(f.second))
        newFuncs.emplace(f);

    _functions = newFuncs;

    return funcSpecs.empty();
  }

  RegisteredFuncPtr FileSource::registerFunc(
    const std::shared_ptr<const FuncSpec>& spec)
  {
    auto& name = spec->name();
    auto iFunc = _functions.find(name);
    if (iFunc != _functions.end())
    {
      auto& ptr = iFunc->second;

      // Attempt to patch the function context to refer to the new function
      auto success = ptr->reregister(spec);
      if (success)
        return ptr;
      
      if (!ptr->deregister())
        return RegisteredFuncPtr();

      _functions.erase(iFunc);
    }

    auto ptr = xloil::registerFunc(spec);
    if (!ptr) 
      return RegisteredFuncPtr();
    _functions.emplace(name, ptr);
    return ptr;
  }

  bool FileSource::deregister(const std::wstring& name)
  {
    auto iFunc = _functions.find(name);
    if (iFunc != _functions.end() && xloil::deregisterFunc(iFunc->second))
    {
      _functions.erase(iFunc);
      return true;
    }
    return false;
  }

  void FileSource::registerLocal(
    const wchar_t * workbookName, 
    const std::vector<std::shared_ptr<const FuncInfo>>& funcInfo, 
    const std::vector<ExcelFuncObject> funcs)
  {
    if (!_workbookName.empty() && _workbookName != workbookName)
      XLO_THROW("Cannot link more than one workbook with the same source");
    xloil::registerLocalFuncs(workbookName, funcInfo, funcs);
    _workbookName = workbookName;
  }

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    FileSource::findFileContext(const wchar_t* source)
  {
    return xloil::findFileSource(source);
  }

  void
    FileSource::deleteFileContext(const std::shared_ptr<FileSource>& source)
  {
    xloil::deleteFileSource(source);
  }

  std::shared_ptr<spdlog::logger> AddinContext::getLogger() const
  {
    return loggerRegistry().default_logger();
  }

  void AddinContext::removeFileSource(ContextMap::const_iterator which)
  {
    _files.erase(which);
  }
}