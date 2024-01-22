#include <xlOil/Interface.h>
#include <xlOil-XLL/FuncRegistry.h>
#include <xlOil-Dynamic/LocalFunctions.h>
#include <xlOil/ObjectCache.h>
#include <xlOilHelpers/Settings.h>
#include <xlOil/Log.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/Loaders/AddinLoader.h>
#include <xlOil/Loaders/PluginLoader.h>
#include <xloil/State.h>
#include <xlOilHelpers/GuidUtils.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/WorkbookScopeFunctions.h>
#include <filesystem>
#include <set>
#include <toml++/toml.h>

using std::make_pair;
using std::wstring;
using std::make_shared;
using std::shared_ptr;
using std::vector;

namespace fs = std::filesystem;

namespace
{
  template<class T, class U>
  std::weak_ptr<T>
    static_pointer_cast(std::weak_ptr<U> const& r)
  {
    return std::static_pointer_cast<T>(std::shared_ptr<U>(r.lock()));
  }

  // {9B8C6F9F-B0FF-46B7-8376-2A6DDECD1B5E}
  static const GUID theXloilNamespace =
    { 0x9b8c6f9f, 0xb0ff, 0x46b7, { 0x83, 0x76, 0x2a, 0x6d, 0xde, 0xcd, 0x1b, 0x5e } };
}

namespace xloil
{
  AddinContext::AddinContext(
    const wchar_t* pathName, const std::shared_ptr<const toml::table>& settings)
    : _settings(settings)
    , _pathName(pathName)
  {
  }

  AddinContext::~AddinContext()
  {}

  void AddinContext::loadPlugins()
  {
    if (!settings())
      return;

    auto addinSettings = (*settings())["Addin"];

    auto pluginNames = Settings::plugins(addinSettings);

    auto plugins = std::set<wstring>(pluginNames.cbegin(), pluginNames.cend());


    const auto xllDir = fs::path(pathName()).remove_filename();
    const auto coreDir = fs::path(Environment::coreDllPath()).remove_filename();

    // If the settings specify a search pattern for plugins, 
    // find the DLLs and add them to our plugins collection
    
    auto searchPattern = Settings::pluginSearchPattern(addinSettings);
    if (!searchPattern.empty())
    {
      WIN32_FIND_DATA fileData;

      auto searchPath = xllDir / searchPattern;
      auto fileHandle = FindFirstFile(searchPath.c_str(), &fileData);
      if (fileHandle != INVALID_HANDLE_VALUE &&
        fileHandle != (void*)ERROR_FILE_NOT_FOUND)
      {
        do
        {
          if (_wcsicmp(fileData.cFileName, Environment::coreDllName()) == 0)
            continue;

          plugins.emplace(fs::path(fileData.cFileName).stem());
        } while (FindNextFile(fileHandle, &fileData));
      }
    }

    for (auto& plugin : plugins)
    {
      if (loadPluginForAddin(*this, plugin))
        _plugins.emplace_back(plugin);
    }
  }

  void AddinContext::detachPlugins()
  {
    for (auto& plugin : _plugins)
      detachPluginForAddin(*this, plugin);
  }

  void FuncSource::init()
  {
    if (!Environment::excelProcess().isEmbedded())
      XLO_THROW("Function registration is only possible when xlOil is running inside Excel");
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

  std::vector<std::shared_ptr<const WorksheetFuncSpec>> FuncSource::functions() const
  {
    vector<shared_ptr<const WorksheetFuncSpec>> result;
    for (auto& item : _functions)
      result.push_back(item.second->spec());
    return result;
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

  FileSource::FileSource(
    const wchar_t* sourcePath, bool watchFile)
    : _sourcePath(sourcePath)
    , _watchFile(watchFile)
  {
    const auto isUrl = wcsncmp(sourcePath, L"http", 4) == 0;
    const auto separator = isUrl ? L'/' : L'\\';
    auto lastSlash = wcsrchr(_sourcePath.c_str(), separator);
    _sourceName = lastSlash ? lastSlash + 1 : _sourcePath.c_str();
    // TODO: implement std::string _functionPrefix;
    //_functionPrefix = toml::find_or<std::string>(*_settings, "FunctionPrefix", "");
  }

  FileSource::~FileSource()
  {
    XLO_DEBUG(L"Deregistering functions in source '{0}'", _sourcePath);
  }

  void FileSource::reload()
  {}

  void FileSource::init()
  {
    if (_watchFile)
    {
      auto path = fs::path(name());
      auto dir = path.remove_filename();
      if (!dir.empty())
        _fileWatcher = Event::DirectoryChange(dir)->weakBind(
          static_pointer_cast<FileSource>(weak_from_this()),
          &FileSource::handleDirChange);
    }

    FuncSource::init();
  }

  LinkedSource::~LinkedSource()
  {
    if (!_localFunctions.empty())
      clearLocalFunctions(_localFunctions);
  }

  void LinkedSource::registerLocal(
    const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcSpecs,
    const bool append)
  {
    if (_workbookName.empty())
      XLO_THROW("Need a linked workbook to declare local functions");

    // Local functions can be registered from any source (e.g. jupyter) but 
    // the linked source (the py file with the same name as the workbook) is
    // special. If this is detected, the VBA module name is set to a special
    // string and all other xlOil local function stubs are removed - we assume
    // that the linked source is loaded before any other.
    auto action = append ? LocalFuncs::APPEND_MODULE : LocalFuncs::REPLACE_MODULE;

    if (_vbaModuleName.empty())
    {
      if (fs::path(filename()).stem() == fs::path(_workbookName).stem())
      {
        _vbaModuleName = wstring(theAutoGenModulePrefix) + L"_linked__";
        action = LocalFuncs::CLEAR_MODULES;
      }
      else
      {
        // Limits: alphanumeric and underscore, 31 chars
        _vbaModuleName = wstring(theAutoGenModulePrefix) + filename();

        auto invalidChars = std::any_of(_vbaModuleName.begin(), _vbaModuleName.end(),
          [](auto c) { return isalnum((int)c) == 0 && c != '.' && c != '_'; });

        if (_vbaModuleName.size() > 31 || invalidChars)
        {
          GUID guid;
          stableGuidFromString(guid, theXloilNamespace, filename());
          _vbaModuleName = wstring(theAutoGenModulePrefix) + guidToWString(guid, GuidToString::BASE62);
        }
        else
          std::replace(_vbaModuleName.begin(), _vbaModuleName.end(), L'.', L'_');
      }
    }

    registerLocalFuncs(
      _localFunctions, 
      _workbookName.c_str(), 
      funcSpecs, 
      _vbaModuleName.c_str(), 
      action);
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
  namespace
  {
    // TODO: is this really necessary or is further refactoring needed?
    void deleteSource(const std::shared_ptr<FuncSource>& context)
    {
      for (auto& [name, addinCtx] : currentAddinContexts())
      {
        auto found = addinCtx->sources().find(context->name());
        if (found != addinCtx->sources().end())
          addinCtx->erase(found->second);
      }
    }
  }
  void LinkedSource::handleClose(const wchar_t* wbName)
  {
    if (_wcsicmp(wbName, linkedWorkbook().c_str()) == 0)
      deleteSource(shared_from_this());
  }

  void LinkedSource::handleRename(const wchar_t* wbName, const wchar_t* prevName)
  {
    if (_wcsicmp(prevName, linkedWorkbook().c_str()) != 0)
      renameWorkbook(wbName);
  }

  void LinkedSource::renameWorkbook(const wchar_t* /*newName*/)
  {
  }

  void FileSource::handleDirChange(
    const wchar_t* /*dirName*/,
    const wchar_t* fileName,
    const Event::FileAction action)
  {
    // Nothing to do if filename does not mach
    if (_wcsicmp(fileName, _sourceName) != 0)
      return;
    
    // TODO: assert check that directory name matches (it should!)
    runExcelThread([
        self = std::static_pointer_cast<FileSource>(shared_from_this()),
        action]()
      {
        switch (action)
        {
          case Event::FileAction::Modified:
          {
            XLO_INFO(L"Module '{0}' modified, reloading.", self->name().c_str());
            self->reload();
            break;
          }
          case Event::FileAction::Delete:
          {
            XLO_INFO(L"Module '{0}' deleted/renamed, removing functions.", self->name().c_str());
            deleteSource(self);
            break;
          }
        }
      }, ExcelRunQueue::ENQUEUE);
  }
}