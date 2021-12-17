#include <xlOil/Interface.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil-Dynamic/LocalFunctions.h>
#include <xlOil/ObjectCache.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Loaders/EntryPoint.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelApp.h>
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
    const wchar_t* linkedWorkbook)
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
    if (_functions.empty() && _workbookName.empty())
      return;

    XLO_DEBUG(L"Deregistering functions in source '{0}'", _sourcePath);

    decltype(_functions) functions;
    wstring workbookName;
    std::swap(_functions, functions);
    std::swap(_workbookName, workbookName);

    if (!workbookName.empty())
      clearLocalFunctions(workbookName.c_str());

    runExcelThread([=]() // TODO: move semanatics rather than copy functions?
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

  void FileSource::registerFuncs(
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

  bool FileSource::deregister(const std::wstring& name)
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

  void FileSource::registerLocal(
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

  std::pair<shared_ptr<FileSource>, shared_ptr<AddinContext>>
    FileSource::findSource(const wchar_t* source)
  {
    auto found = xloil::findFileSource(source);
    if (found.first)
    {
      // Slightly gross little check that the linked workbook is still open
      // Can we do better?
      const auto& wbName = found.first->_workbookName;
      if (!wbName.empty() && !COM::checkWorkbookIsOpen(wbName.c_str()))
      {
        deleteSource(found.first);
        return make_pair(shared_ptr<FileSource>(), shared_ptr<AddinContext>());
      }
    }
    return found;
  }

  void FileSource::deleteSource(const shared_ptr<FileSource>& source)
  {
    xloil::deleteFileSource(source);
  }

  void FileSource::startWatch()
  {
    if (!linkedWorkbook().empty())
      _workbookWatcher = Event::WorkbookAfterClose().weakBind(weak_from_this(), &FileSource::handleClose);
  }

  void WatchedSource::reload()
  {}

  void WatchedSource::startWatch()
  {
    auto path = fs::path(sourcePath());
    auto dir = path.remove_filename();
    if (!dir.empty())
      _fileWatcher = Event::DirectoryChange(dir)->weakBind(weak_from_this(), &WatchedSource::handleDirChange);
  }


  void FileSource::handleClose(const wchar_t* wbName)
  {
    if (_wcsicmp(wbName, linkedWorkbook().c_str()) == 0)
      FileSource::deleteSource(shared_from_this());
  }

  void WatchedSource::handleDirChange(
    const wchar_t* dirName,
    const wchar_t* fileName,
    const Event::FileAction action)
  {
    if (_wcsicmp(fileName, sourceName()) != 0)
      return;

    runExcelThread([
      self = std::static_pointer_cast<WatchedSource>(shared_from_this()),
        filePath = fs::path(dirName) / fileName,
        action]()
      {
        // File paths should match as our directory watch listener only checks
        // the specified directory
        assert(_wcsicmp(filePath.c_str(), self->sourcePath().c_str()) == 0);

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
          FileSource::deleteSource(self);
          break;
        }
        }
      }, ExcelRunQueue::ENQUEUE);
  }

  void AddinContext::removeSource(ContextMap::const_iterator which)
  {
    _files.erase(which);
  }
}