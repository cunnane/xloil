#include <xlOil/Interface.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil/Register/LocalFunctions.h>
#include <xlOil/ObjectCache.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/Log.h>
#include <xlOil/ApiCall.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xloil/State.h>
#include <xlOil-COM/Connect.h>

using std::make_pair;
using std::wstring;
using std::make_shared;
using std::shared_ptr;

namespace xloil
{
  AddinContext::AddinContext(
    const wchar_t* pathName, std::shared_ptr<const toml::table> settings)
    : _settings(settings)
    , _pathName(pathName)
  {
  }

  AddinContext::~AddinContext()
  {
  }

  FileSource::FileSource(
    const wchar_t* sourcePath, 
    const wchar_t* linkedWorkbook,
    bool watchSource)
    : _sourcePath(sourcePath)
  {
    auto lastSlash = wcsrchr(_sourcePath.c_str(), L'\\');
    _sourceName = lastSlash ? lastSlash + 1 : _sourcePath.c_str();
    if (linkedWorkbook)
      _workbookName = linkedWorkbook;
    //_functionPrefix = toml::find_or<std::string>(*_settings, "FunctionPrefix", "");
  }

  FileSource::~FileSource()
  {
    excelApiCall([this]()
    {
      XLO_DEBUG(L"Deregistering functions in source '{0}'", _sourcePath);
      forgetLocalFunctions(_workbookName.c_str());
      for (auto& f : _functions)
        xloil::deregisterFunc(f.second);
      _functions.clear();
    }, QueueType::XLL_API);
  }

  void FileSource::registerFuncs(
    const std::vector<std::shared_ptr<const FuncSpec> >& funcSpecs)
  {
    excelApiCall([specs = funcSpecs, self = this]() mutable
    {
      decltype(self->_functions) newFuncs;

      for (auto& f : specs)
      {
        // If registration succeeds, just add the function to the new map
        auto ptr = self->registerFunc(f);
        if (ptr)
        {
          self->_functions.erase(f->name());
          newFuncs.emplace(f->name(), ptr);
          f.reset();
        }
      }

      // Remove all the null FuncSpec ptrs
      specs.erase(
        std::remove_if(specs.begin(), specs.end(), [](auto& f) { return !f; }),
        specs.end());

      // Any functions remaining in the old map must have been removed from the module
      // so we can deregister them, but if that fails we have to keep them or they
      // will be orphaned
      for (auto& f : self->_functions)
        if (!xloil::deregisterFunc(f.second))
          newFuncs.emplace(f);

      self->_functions = newFuncs;

      for (auto& f : specs)
        XLO_ERROR(L"Registration failed for: {0}", f->name());

    }, QueueType::XLL_API);
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
    if (iFunc != _functions.end())
    {
      excelApiCall([iFunc, self = this]()
      {
        if (xloil::deregisterFunc(iFunc->second))
          self->_functions.erase(iFunc);
      }, QueueType::XLL_API);
      return true;
    }
    return false;
  }

  void FileSource::registerLocal(
    const std::vector<std::shared_ptr<const FuncInfo>>& funcInfo, 
    const std::vector<ExcelFuncObject> funcs)
  {
    if (_workbookName.empty())
      XLO_THROW("Need a linked workbook to declare local functions");
    excelApiCall([=, self = this]()
    {
      xloil::registerLocalFuncs(self->_workbookName.c_str(), funcInfo, funcs);
    });
  }

  std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
    FileSource::findFileContext(const wchar_t* source)
  {
    auto found = xloil::findFileSource(source);
    if (found.first)
    {
      // Slightly gross little check that the linked workbook is still open
      // Can we do better?
      const auto& wbName = found.first->_workbookName;
      if (!wbName.empty() && !COM::checkWorkbookIsOpen(wbName.c_str()))
      {
        deleteFileContext(found.first);
        return make_pair(shared_ptr<FileSource>(), shared_ptr<AddinContext>());
      }
    }
    return found;
  }

  void
    FileSource::deleteFileContext(const std::shared_ptr<FileSource>& source)
  {
    xloil::deleteFileSource(source);
  }

  void 
    AddinContext::removeSource(ContextMap::const_iterator which)
  {
    _files.erase(which);
  }
}