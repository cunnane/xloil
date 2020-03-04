#pragma once
#include "ExportMacro.h"
#include "Register.h"
#include "Events.h"
#include "ExcelObj.h"
#include "ExcelObjCache.h"
#include <memory>
#include <map>

namespace toml {
  template<typename, template<typename...> class, template<typename...> class> class basic_value;
  struct discard_comments;
  using value = basic_value<discard_comments, std::unordered_map, std::vector>;
}
namespace Excel { struct _Application; }
namespace xloil { class RegisteredFunc; }

namespace xloil
{

  class XLOIL_EXPORT Core
  {
  public:
    Core(const wchar_t* pluginName, std::shared_ptr<const toml::value> settings);
    ~Core();

    /// <summary>
    /// Returns the full path to the xloil Core dll, including the filename
    /// </summary>
    /// <returns></returns>
    static const wchar_t* theCorePath();

    /// <summary>
    /// Returns just the filename of the xloil Core dll
    /// </summary>
    /// <returns></returns>
    static const wchar_t* theCoreName();

    static Excel::_Application& theExcelApp();

    /// <summary>
    /// See templated version <see cref="registerFunc"/>
    /// </summary>
    /// <param name="info"></param>
    /// <param name="callback"></param>
    /// <param name="data"></param>
    /// <param name="group"></param>
    /// <returns></returns>
    int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info, 
        RegisterCallback callback, 
        const std::shared_ptr<void>&  data) noexcept;

    /// <summary>
    /// Probably shouldn't use this
    /// </summary>
    /// <param name="info"></param>
    /// <param name="callback"></param>
    /// <param name="data"></param>
    /// <param name="group"></param>
    /// <returns></returns>
    int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info,
        AsyncCallback callback,
        const std::shared_ptr<void>& data) noexcept;

    /// <summary>
    /// Registers an exported function
    /// </summary>
    /// <param name="info"></param>
    /// <param name="functionName">The mangled (if not extern C) entry point</param>
    /// <returns>The registration ID or zero on failure</returns>
    int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info, 
        const char* functionName) noexcept;

    /// <summary>
    /// Registers a function object with suitable signature as an
    /// Excel function
    /// </summary>
    /// <param name="info"></param>
    /// <param name="f"></param>
    /// <returns>The registration ID or zero on failure</returns>
    int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info, 
        const ExcelFuncPrototype& f) noexcept;

    /// <summary>
    /// 
    /// </summary>
    /// <param name="info">Pointer to info which determines how the function appears in Excel</param>
    /// <param name="callback">Pointer to callback function with correct signature</param>
    /// <param name="data">Pointer to context data that will be returned with the callback</param>
    /// </param>
    /// <returns>The registration ID produced by Excel which can be used to invoke the
    ///  funcion The registration ID or zero on failure</returns>
    template <class TData> inline int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info, 
        RegisterCallbackT<TData> callback,
        const std::shared_ptr<TData>& data) noexcept
    {
      return registerFunc(info, (RegisterCallback)callback, std::static_pointer_cast<void>(data));
    }


    template <class TData> inline int
      registerFunc(
        const std::shared_ptr<const FuncInfo>& info,
        AsyncCallbackT<TData> callback,
        const std::shared_ptr<TData>& data) noexcept
    {
      return registerFunc(info, (AsyncCallback)callback, std::static_pointer_cast<void>(data));
    }

    /// <summary>
    /// Searches existing functions for one with a name matching that
    /// in the info. If not found, returns false. If found, attempts
    /// to patch in the provided context and any changes to help strings 
    /// in the info.  If the changes are too significant, e.g. the number
    /// of arguments changed, it deregisters the found function and returns 
    /// false.
    /// </summary>
    /// <param name="info"></param>
    /// <param name="newContext"></param>
    /// <returns>false if the function was not found or patching failed</returns>
    bool 
      reregister(
        const std::shared_ptr<const FuncInfo>& info,
        const std::shared_ptr<void>& newContext);

    /// <summary>
    /// Removes the specified function from Excel
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    bool
      deregister(const std::wstring& name);

    void
      deregisterAll();

    /// <summary>
    /// Returns true if the provided string contains the magic chars
    /// for the ExcelObj cache. Expects a counted string.
    /// </summary>
    /// <param name="str">Pointer to string start</param>
    /// <param name="length">Number of chars to read</param>
    /// <returns></returns>
    static inline bool 
      maybeCacheReference(const wchar_t* str, size_t length)
    {
      return checkObjectCacheReference(str, length);
    }

    static bool
      fetchCache(const wchar_t* cacheString, size_t length, std::shared_ptr<const ExcelObj>& obj);

    static ExcelObj
      insertCache(std::shared_ptr<const ExcelObj>&& obj);

    static inline ExcelObj
      insertCache(ExcelObj&& obj)
    {
      return insertCache(std::make_shared<const ExcelObj>(obj));
    }

    std::shared_ptr<spdlog::logger> getLogger()
    {
      return loggerRegistry().default_logger();
    }

    const toml::value* settings() const { return _settings.get(); }
    const std::wstring& pluginName() const { return _pluginName; }

  private:
    const std::wstring _pluginName;
    std::string _functionPrefix;
    std::shared_ptr<const toml::value> _settings;
    std::map<std::wstring, std::shared_ptr<RegisteredFunc>> _functions;
  };

#define XLO_PLUGIN_INIT_FUNC xloil_init
#define XLO_PLUGIN_EXIT_FUNC xloil_exit

#define XLO_PLUGIN_INIT(...) extern "C" __declspec(dllexport) int XLO_PLUGIN_INIT_FUNC##(__VA_ARGS__) noexcept
#define XLO_PLUGIN_EXIT() extern "C" __declspec(dllexport) int XLO_PLUGIN_EXIT_FUNC##() noexcept

  typedef int(*pluginInitFunc)(Core&);
  typedef int(*pluginExitFunc)();
}
