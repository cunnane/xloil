#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/Register.h>
#include <xloil/FuncSpec.h>
#include <xloil/Version.h>
#include <memory>
#include <map>

namespace toml { inline namespace v3 { class table; } }
namespace xloil { 
  class RegisteredWorksheetFunc; 
  class LocalWorksheetFunc;
  class AddinContext; 
  namespace Event { enum class FileAction; }
}

namespace xloil
{
  class AddinContext;
  /// <summary>
  /// FuncSource is the base class in a hierachy which keeps track 
  /// of a location (typically a file) from which UDFs have been
  /// registered and provides functionality such as reloading on change.
  /// 
  /// Plugins should avoid keeping references to a FuncSource, or if
  /// they do be careful to clean them up when an XLL detaches
  /// </summary>
  class XLOIL_EXPORT FuncSource : public std::enable_shared_from_this<FuncSource>
  {
  public:
    /// <summary>
    /// Called just after the source is registered in the global FuncSource
    /// map. FuncSources should not register functions before `init`.
    /// </summary>
    virtual void init() = 0;
  
    virtual ~FuncSource();

    virtual const std::wstring& name() const = 0;

    /// <summary>
    /// Registers the given function specifcations with Excel. If
    /// registration fails the input parameter will contain the failed
    /// functions, otherwise it will be empty. 
    /// 
    /// If this function is called a second time it replaces all currently
    /// registered functions unless append is true.
    /// 
    /// </summary>
    /// <param name="specs">functions to register</param>
    /// <param name="append">if false, replacess all currently registered functions</param>
    void registerFuncs(
      const std::vector<std::shared_ptr<const WorksheetFuncSpec> >& specs,
      const bool append = false);

    /// <summary>
    /// Removes the specified function from Excel
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    bool deregister(const std::wstring& name);

    /// <summary>
    /// A list of all functions declared by this Source.
    /// </summary>
    std::vector<std::shared_ptr<const WorksheetFuncSpec>> functions() const;

  private:
    std::map<std::wstring, std::shared_ptr<RegisteredWorksheetFunc>> _functions;
  };

  /// <summary>
  /// FileSource extends FuncSource by watching the specified source file
  /// for changes and signals `reload` on modification or deletes the 
  /// source (hence unregistering the UDFs) if the watched file is removed.
  /// </summary>
  class XLOIL_EXPORT FileSource : public FuncSource
  {
  public:
    friend class AddinContext;

    /// <summary>
    /// </summary>
    /// <param name="sourcePath">Should be a full pathname</param>
    /// <param name="watchFile">If true enables a file watch</param>
    FileSource(const wchar_t* sourcePath, bool watchFile = false);
    virtual ~FileSource();

    virtual const std::wstring& name() const override { return _sourcePath; }
    const wchar_t* filename() const { return _sourceName; }

  protected:
    /// <summary>
    /// Invoked when the source file is modified, but not deleted.
    /// </summary>
    virtual void reload();
    virtual void init();

  private:
    std::wstring _sourcePath;
    const wchar_t* _sourceName;
    bool _watchFile;
    std::shared_ptr<const void> _fileWatcher;

    void handleDirChange(
      const wchar_t* dirName,
      const wchar_t* fileName,
      const Event::FileAction action);
  };

  /// <summary>
  /// LinkedSource extends FileSource by having an associated workbook. 
  /// This allows the registration of local functions. The workbook is 
  /// watched for rename and close events which trigger a call of 
  /// `renameWorkbook` or deletion of the source respectively.
  /// </summary>
  class XLOIL_EXPORT LinkedSource : public FileSource
  {
  public:
    LinkedSource(
      const wchar_t* sourceName,
      bool watchFile,
      const wchar_t* linkedWorkbookName)
      : FileSource(sourceName, watchFile)
    {
      if (linkedWorkbookName)
        _workbookName = linkedWorkbookName;
    }
    ~LinkedSource();

    /// <summary>
    /// Registers the given functions as local functions in the specified
    /// workbook.  Either appends to or overwrites the existing functions
    /// depending on the 'append' parameter.
    /// </summary>
    void registerLocal(
      const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcSpecs, 
      const bool append);

    const std::wstring& linkedWorkbook() const { return _workbookName; }

  protected:
    virtual void init();
    virtual void renameWorkbook(const wchar_t* newPathName);

  private:
    std::map<std::wstring, std::shared_ptr<const LocalWorksheetFunc>> _localFunctions;
    std::wstring _workbookName;
    std::wstring _vbaModuleName;
    std::shared_ptr<const void> _workbookCloseHandler;
    std::shared_ptr<const void> _workbookRenameHandler;

    void handleClose(const wchar_t* wbName);
    void handleRename(const wchar_t* wbPathName, const wchar_t* prevName);
  };

  /// <summary> 
  /// The AddinContext keeps track of file sources associated with an Addin
  /// to ensure they are properly cleaned up when the addin unloads
  /// </summary>
  class AddinContext
  {
  public:
    using ContextMap = std::map<std::wstring, std::shared_ptr<FuncSource>>;

    AddinContext(
      const wchar_t* pathName, 
      const std::shared_ptr<const toml::table>& settings);

    ~AddinContext();


    /// <summary>
    /// Looks for a FileSource corresponding the specified pathname.
    /// Returns the FileSource if found, otherwise a null pointer
    /// </summary>
    /// <param name="sourcePath"></param>
    /// <returns></returns>
    XLOIL_EXPORT static std::pair<std::shared_ptr<FuncSource>, std::shared_ptr<AddinContext>>
      findSource(const wchar_t* sourcePath);

    /// <summary>
    /// Removes the specified source
    /// </summary>
    void erase(const std::shared_ptr<FuncSource>& context)
    {
      _files.erase(context->name());
    }

    /// <summary>
    /// Gets the root of the addin's ini file
    /// </summary>
    auto settings() const { return _settings.get(); }

    /// <summary>
    /// Returns a map of all Sources associated with this XLL addin
    /// </summary>
    const ContextMap& sources() const { return _files; }

    /// <summary>
    /// Returns the full pathname of the XLL addin
    /// </summary>
    const std::wstring& pathName() const { return _pathName; }

    /// <summary>
    /// Returns the filename of the XLL addin
    /// </summary>
    const wchar_t* fileName() const 
    {
      auto slash = _pathName.find_last_of(L'\\');
      return _pathName.c_str() + slash + 1;
    }

    void addSource(const std::shared_ptr<FuncSource>& source)
    {
      _files.emplace(std::make_pair(source->name(), source));
      source->init();
    }

    std::wstring logFilePath;

  private:
    AddinContext(const AddinContext&) = delete;
    AddinContext& operator=(const AddinContext&) = delete;

    std::wstring _pathName;
    std::shared_ptr<const toml::table> _settings;
    ContextMap _files;
  };

/// <summary>
/// This macro declares the plugin entry point. Its arguments must match
/// <see cref="PluginInitFunc"/>.
/// </summary>
#define XLO_PLUGIN_INIT(...) extern "C" __declspec(dllexport) int \
  XLO_PLUGIN_INIT_FUNC##(__VA_ARGS__) noexcept

#define XLO_PLUGIN_INIT_FUNC xloil_init

  /// <summary>
  /// Contains information the plugin can use to initialise or 
  /// link with another addin
  /// </summary>
  /// 
  // TODO: rename to maybe PluginAction?  Also maybe Extensions rather than plugin?
  struct PluginContext
  {
    enum Action
    {
      /// <summary>
      /// The Load action is specified the first time a plugin is initialised
      /// </summary>
      Load,
      /// <summary>
      /// The Attach action is used when an XLL addin has requested use of the 
      /// plugin. The addin may have a settings file which the plugin should process
      /// </summary>
      Attach,
      /// <summary>
      /// Detach is called when an XLL using the plugin is unloading
      /// </summary>
      Detach,
      /// <summary>
      /// When Unload is called, the plugin should clean up all internal
      /// data in anticipation of a call to FreeLibrary.
      /// </summary>
      Unload
    };
    Action action;
    const wchar_t* pluginName;
    const toml::table& settings;
    uint8_t coreMajorVersion;
    uint8_t coreMinorVersion;
    uint8_t corePatchVersion;

    bool checkExactVersion() const
    {
      return coreMajorVersion == XLOIL_MAJOR_VERSION 
        && coreMinorVersion == XLOIL_MINOR_VERSION 
        && corePatchVersion == XLOIL_PATCH_VERSION;
    }
  };

  /// <summary>
  /// A plugin must declare an extern C function with this signature
  /// </summary>
  typedef int(*PluginInitFunc)(AddinContext*, const PluginContext&);
}
