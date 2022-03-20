#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/Register.h>
#include <xloil/FuncSpec.h>
#include <xloil/Version.h>
#include <memory>
#include <map>

namespace toml { class table; }
namespace xloil { 
  class RegisteredWorksheetFunc; 
  class AddinContext; 
  namespace Event { enum class FileAction; }
}

namespace xloil
{
  class AddinContext;

  /// <summary>
  /// A file source collects Excel UDFs created from a single file.
  /// The file could be a plugin DLL or source file. You can inherit
  /// from this class to provide additional tracking functionality.
  /// 
  /// Plugins should avoid keeping references to file sources, or if
  /// they do be careful to clean them up when an XLL detaches
  /// </summary>
  class XLOIL_EXPORT FileSource : public std::enable_shared_from_this<FileSource>
  {
  public:
    /// <summary>
    /// 
    /// </summary>
    /// <param name="sourcePath">Should be a full pathname</param>
    /// <param name="linkedWorkbook">Name of linked workbook, required for local functions</param>
    FileSource(
      const wchar_t* sourcePath,
      const wchar_t* linkedWorkbook = nullptr);
    virtual ~FileSource();

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
    /// Registers the given functions as local functions in the specified
    /// workbook
    /// </summary>
    /// <param name="funcSpecs"></param>
    /// <param name="append">if false, replacess all currently registered functions</param>
    void registerLocal(
        const std::vector<std::shared_ptr<const WorksheetFuncSpec>>& funcSpecs,
        const bool append);

    /// <summary>
    /// Looks for a FileSource corresponding the specified pathname.
    /// Returns the FileSource if found, otherwise a null pointer
    /// </summary>
    /// <param name="sourcePath"></param>
    /// <returns></returns>
    static std::pair<std::shared_ptr<FileSource>, std::shared_ptr<AddinContext>>
      findSource(const wchar_t* sourcePath);

    /// <summary>
    /// Removes the specified source from all add-in contexts
    /// </summary>
    /// <param name="context"></param>
    static void
      deleteSource(const std::shared_ptr<FileSource>& context);

    const std::wstring& sourcePath() const { return _sourcePath; }
    const std::wstring& linkedWorkbook() const { return _workbookName; }
    const wchar_t* sourceName() const { return _sourceName; }

  protected:
    friend class AddinContext;
    virtual void startWatch();

  private:
    std::map<std::wstring, std::shared_ptr<RegisteredWorksheetFunc>> _functions;
    std::wstring _sourcePath;
    const wchar_t* _sourceName;
    std::wstring _workbookName;

    std::shared_ptr<const void> _workbookWatcher;

    void handleClose(const wchar_t* wbName);
    // TODO: implement std::string _functionPrefix;
  };;

  class WatchedSource : public FileSource
  {
  public:
    WatchedSource(
      const wchar_t* sourceName,
      const wchar_t* linkedWorkbook = nullptr)
      : FileSource(sourceName, linkedWorkbook)
    {}

  protected:
    XLOIL_EXPORT virtual void startWatch();
    /// <summary>
    /// Invoked when the watched file is modified, but not deleted
    /// </summary>
    virtual void reload() = 0;
  private:
    std::shared_ptr<const void> _fileWatcher;

    void handleDirChange(
      const wchar_t* dirName,
      const wchar_t* fileName,
      const Event::FileAction action);
  };

  /// <summary> 
  /// The AddinContext keeps track of file sources associated with an Addin
  /// to ensure they are properly cleaned up when the addin unloads
  /// </summary>
  class AddinContext
  {
  public:
    using ContextMap = std::map<std::wstring, std::shared_ptr<FileSource>>;

    AddinContext(
      const wchar_t* pathName, 
      const std::shared_ptr<const toml::table>& settings);

    ~AddinContext();

    /// <summary>
    /// Links a FileSource for the specified source path to this
    /// add-in context. Other addin contexts are first searched
    /// for the matching FileSource.  If it is not found a new
    /// one is created passing the variadic argument to the TSource
    /// constructor.
    /// </summary>
    template <class TSource, class...Args>
    std::pair<std::shared_ptr<TSource>, bool>
      tryAdd(
        const wchar_t* sourcePath, Args&&... args)
    {
      auto found = FileSource::findSource(sourcePath).first;
      if (found)
      {
        addSource(found, false);
        return std::make_pair(std::static_pointer_cast<TSource>(found), false);
      }
      else
      {
        auto newSource = std::make_shared<TSource>(std::forward<Args>(args)...);
        addSource(newSource);
        return std::make_pair(newSource, true);
      }
    }

    /// <summary>
    /// Gets the root of the addin's ini file
    /// </summary>
    const toml::table* settings() const { return _settings.get(); }

    /// <summary>
    /// Returns a map of all FileSource associated with this XLL addin
    /// </summary>
    const ContextMap& files() const { return _files; }

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

    void addSource(const std::shared_ptr<FileSource>& source)
    {
      _files.emplace(std::make_pair(source->sourcePath(), source));
      source->startWatch();
    }

    void removeSource(ContextMap::const_iterator which);

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
