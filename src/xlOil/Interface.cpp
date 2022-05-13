#include <xlOil/Interface.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil-Dynamic/LocalFunctions.h>
#include <xlOil/ObjectCache.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xloil/State.h>
#include <xlOil-COM/Connect.h>
#include <filesystem>
using std::make_pair;
using std::wstring;
using std::make_shared;
using std::shared_ptr;
using std::vector;

namespace fs = std::filesystem;

namespace xloil
{
  AddinContext::AddinContext(
    const wchar_t* pathName, const std::shared_ptr<const toml::table>& settings)
    : _settings(settings)
    , _pathName(pathName)
  {
  }

  AddinContext::~AddinContext()
  {
  }

  void FuncSource::init()
  {
  }

  FuncSource::~FuncSource()
  {
    if (_functions.empty())
      return;

    decltype(_functions) functions;
    std::swap(_functions, functions);

    runExcelThread([functions = std::move(functions)]()
    {
      for (auto& f : functions)
        f.second->deregister();
    }, ExcelRunQueue::XLL_API);
  }

  namespace
  {
    auto registerFunc(
      std::map<std::wstring, std::shared_ptr<RegisteredWorksheetFunc>>& existingFuncs,
      const shared_ptr<const WorksheetFuncSpec>& spec)
    {
      auto& name = spec->name();
      auto iFunc = existingFuncs.find(name);
      if (iFunc != existingFuncs.end())
      {
        auto& ptr = iFunc->second;

        // Attempt to patch the function context to refer to the new function
        if (ptr->reregister(spec))
          return make_pair(ptr, true);

        if (!ptr->deregister())
          return make_pair(ptr, false);

        existingFuncs.erase(iFunc);
      }

      auto ptr = xloil::registerFunc(spec);
      return make_pair(ptr, !!ptr);
    }
  }

  void FuncSource::registerFuncs(
    const std::vector<std::shared_ptr<const WorksheetFuncSpec> >& funcSpecs,
    const bool append)
  {
    runExcelThread([append, specs = funcSpecs, self = shared_from_this()]() mutable
    {
      auto& existingFuncs = self->_functions;
      decltype(self->_functions) newFuncs;

      for (auto& f : specs)
      {
        // If registration succeeds, just add the function to the new map
        auto [ptr, success] = registerFunc(existingFuncs, f);

        // If deregistration fails we have to keep the ptr or it will be orphaned
        if (ptr)
          newFuncs.emplace(f->name(), ptr);

        if (success)
          f.reset(); // Clear pointer in specs to indicate success
      }

      for (auto& f : specs)
        if (f)
          XLO_ERROR(L"Registration failed for: {0}", f->name());

      if (append)
        newFuncs.merge(existingFuncs);
      self->_functions = newFuncs;
    }, ExcelRunQueue::XLL_API);
  }

  bool FuncSource::deregister(const std::wstring& name)
  {
    auto iFunc = _functions.find(name);
    if (iFunc != _functions.end())
    {
      runExcelThread([iFunc, self = this]()
      {
        if (iFunc->second->deregister())
          self->_functions.erase(iFunc);
      }, ExcelRunQueue::XLL_API);
      return true;
    }
    return false;
  }


  std::pair<std::shared_ptr<FuncSource>, std::shared_ptr<AddinContext>>
    AddinContext::findSource(const wchar_t* source)
  {
    for (auto& [addinName, addin] : currentAddinContexts())
    {
      auto found = addin->sources().find(source);
      if (found != addin->sources().end())
        return make_pair(found->second, addin);
    }
    return make_pair(shared_ptr<FuncSource>(), shared_ptr<AddinContext>());
  }

  void
    AddinContext::deleteSource(const std::shared_ptr<FuncSource>& context)
  {
    for (auto& [name, addinCtx] : currentAddinContexts())
    {
      auto found = addinCtx->sources().find(context->name());
      if (found != addinCtx->sources().end())
        addinCtx->_files.erase(found);
    }
  }

  //std::pair<shared_ptr<FuncSource>, shared_ptr<AddinContext>>
  //  findSource(const wchar_t* source)
  //{
  //  auto found = xloil::findFileSource(source);
  //  //if (found.first)
  //  //{
  //  //  // Slightly gross little check that the linked workbook is still open
  //  //  // Can we do better?
  //  //  const auto& wbName = found.first->_workbookName;
  //  //  if (!wbName.empty() && !COM::checkWorkbookIsOpen(wbName.c_str()))
  //  //  {
  //  //    deleteSource(found.first);
  //  //    return make_pair(shared_ptr<FileSource>(), shared_ptr<AddinContext>());
  //  //  }
  //  //}
  //  return found;
  //}

  FileSource::FileSource(
    const wchar_t* sourcePath, bool watchFile)
    : _sourcePath(sourcePath)
  {
    auto lastSlash = wcsrchr(_sourcePath.c_str(), L'\\');
    _sourceName = lastSlash ? lastSlash + 1 : _sourcePath.c_str();
    //_functionPrefix = toml::find_or<std::string>(*_settings, "FunctionPrefix", "");
  }

  FileSource::~FileSource()
  {
    XLO_DEBUG(L"Deregistering functions in source '{0}'", _sourcePath);
  }

  LinkedSource::~LinkedSource()
  {
    wstring workbookName;
    std::swap(_workbookName, workbookName);
    if (!workbookName.empty())
      clearLocalFunctions(workbookName.c_str());
  }

  void LinkedSource::registerLocal(
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcSpecs,
    const bool append)
  {
    if (_workbookName.empty())
      XLO_THROW("Need a linked workbook to declare local functions");
    runExcelThread([=, self = this]()
    {
      xloil::registerLocalFuncs(self->_workbookName.c_str(), funcSpecs, append);
    });
  }

  template<class T, class U>
  std::weak_ptr<T>
    static_pointer_cast(std::weak_ptr<U> const& r)
  {
    return std::static_pointer_cast<T>(std::shared_ptr<U>(r.lock()));
  }


  void FileSource::reload()
  {}

  void FileSource::init()
  {
    auto path = fs::path(name());
    auto dir = path.remove_filename();
    if (!dir.empty())
      _fileWatcher = Event::DirectoryChange(dir)->weakBind(
        static_pointer_cast<FileSource>(weak_from_this()),
        &FileSource::handleDirChange);

    FuncSource::init();
  }

  void LinkedSource::init()
  {
    if (!linkedWorkbook().empty())
    {
      _workbookCloseHandler = Event::WorkbookAfterClose().weakBind(
        static_pointer_cast<LinkedSource>(weak_from_this()),
        &LinkedSource::handleClose);
      _workbookRenameHandler = Event::WorkbookRename().weakBind(
        static_pointer_cast<LinkedSource>(weak_from_this()),
        &LinkedSource::handleRename);

    }
    FileSource::init();
  }

  void LinkedSource::handleClose(const wchar_t* wbName)
  {
    if (_wcsicmp(wbName, linkedWorkbook().c_str()) == 0)
      AddinContext::deleteSource(shared_from_this());
  }

  void LinkedSource::handleRename(const wchar_t* wbName, const wchar_t* prevName)
  {
    if (_wcsicmp(prevName, linkedWorkbook().c_str()) != 0)
      renameWorkbook(wbName);
  }

  void LinkedSource::renameWorkbook(const wchar_t* newName)
  {
    //deleteSource(shared_from_this());
    //// if it's in the same directory and the filename matches...
    //fs::copy_file(sourcePath(), wbName....);
    //rename();
    //createSource()
  }

  void FileSource::handleDirChange(
    const wchar_t* dirName,
    const wchar_t* fileName,
    const Event::FileAction action)
  {
    if (fs::path(fileName) != fs::path(name()))
      return;
    
    runExcelThread([
      self = std::static_pointer_cast<FileSource>(shared_from_this()),
        filePath = fs::path(dirName) / fileName,
        action]()
      {
        switch (action)
        {
          case Event::FileAction::Modified:
          {
            XLO_INFO(L"Module '{0}' modified, reloading.", filePath.c_str());
            self->reload();
            break;
          }
          case Event::FileAction::Delete:
          {
            XLO_INFO(L"Module '{0}' deleted/renamed, removing functions.", filePath.c_str());
            AddinContext::deleteSource(self);
            break;
          }
        }
      }, ExcelRunQueue::ENQUEUE);
  }
}