#include "Interface.h"
#include "internal/FuncRegistry.h"
#include "ObjectCache.h"
#include "Settings.h"
#include "EntryPoint.h"
#include <ComInterface/Connect.h>
#include <toml11/toml.hpp>

using std::make_pair;
using std::wstring;

namespace xloil
{
  Core::Core(const wchar_t* pluginName)
    : _pluginName(pluginName)
    , _settings(fetchPluginSettings(pluginName))
  {
    if (_settings)
    {
      _functionPrefix = toml::find_or<std::string>(*_settings, "FunctionPrefix", "");
    }
    // This collects all statically declared Excel functions, i.e. raw C functions
    // It assumes that this ctor and hence processRegistryQueue is run after each
    // plugin has been loaded, so that all functions on the queue belong to the 
    // current plugin
    for (auto& f : processRegistryQueue(pluginName))
      _functions.emplace(f->info()->name, f);
  }
  Core::~Core()
  {
    deregisterAll();
  }

  const wchar_t* Core::theCorePath()
  {
    return xloil::theCorePath();
  }
  const wchar_t* Core::theCoreName()
  {
    return xloil::theCoreName();
  }

  Excel::_Application & Core::theExcelApp()
  {
    return excelApp();
  }

  int
    Core::registerFunc(const std::shared_ptr<const FuncInfo>& info, RegisterCallback callback, 
      const std::shared_ptr<void>& data) noexcept
  {
    auto ptr = xloil::registerFunc(info, callback, data);
    if (!ptr) return 0;
    _functions.emplace(info->name, ptr);
    return ptr->registerId();
  }

  int
    Core::registerFunc(const std::shared_ptr<const FuncInfo>& info, AsyncCallback callback,
      const std::shared_ptr<void>& data) noexcept
  {
    auto ptr = xloil::registerFunc(info, callback, data);
    if (!ptr) return 0;
    _functions.emplace(info->name, ptr);
    return ptr->registerId();
  }

  int
    Core::registerFunc(const std::shared_ptr<const FuncInfo>& info, const char* functionName) noexcept
  {
    auto ptr = xloil::registerFunc(info, functionName, _pluginName.c_str());
    if (!ptr) return 0;
    _functions.emplace(info->name, ptr);
    return ptr->registerId();
  }

  int Core::registerFunc(const std::shared_ptr<const FuncInfo>& info, const ExcelFuncPrototype & f) noexcept
  {
    auto ptr = xloil::registerFunc(info, f);
    if (!ptr) return 0;
    _functions.emplace(info->name, ptr);
    return ptr->registerId();
  }

  bool Core::reregister(
    const std::shared_ptr<const FuncInfo>& info,
    const std::shared_ptr<void>& newContext)
  {
    auto iFunc = _functions.find(info->name);
    if (iFunc == _functions.end())
      return false;
    auto[name, ptr] = *iFunc;
    auto success = ptr->reregister(info, newContext);
    if (!success)
      ptr->deregister();
    return success;
  }

  bool Core::deregister(const std::wstring& name)
  {
    auto iFunc = _functions.find(name);
    if (iFunc == _functions.end())
      return false;
    xloil::deregisterFunc(iFunc->second);
    _functions.erase(iFunc);
    return true;
  }

  void
    Core::deregisterAll()
  {
    for (auto& f : _functions)
      xloil::deregisterFunc(f.second);
    _functions.clear();
  }
 
  bool
    Core::fetchCache(const wchar_t* cacheString, size_t length, std::shared_ptr<const ExcelObj>& obj)
  {
    return xloil::fetchCacheObject(cacheString, length, obj);
  }

  ExcelObj
    Core::insertCache(const std::shared_ptr<const ExcelObj>& obj)
  {
    return xloil::addCacheObject(obj);
  }
}