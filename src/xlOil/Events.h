#pragma once
#include "ExportMacro.h"
#include <functional>
#include <memory>
#include <list>
#include <mutex>
#include <future>
#include <string>

namespace xloil
{
  struct AsyncVoidCollector
  {
    template<class THandlers, class... Args>
    void operator()(const THandlers& handlers, Args&&... args) const
    {
      (void)std::async([handlers](Args... vals)
      {
        for (auto& h : handlers)
          h(vals...);
      }, args...);
    }
  };

  struct VoidCollector
  {
    template<class THandlers, class... Args>
    void operator()(const THandlers& handlers, Args&&... args) const
    {
      for (auto& h : handlers)
        h(args...);
    }
  };

  template<class, class> class Event {};
  template<class R, class TCollector, class... Args>
  class Event<R(Args...), TCollector>
  {
  public:
    using handler = std::function < R(Args...)>;
    using handler_id = const handler*;

    Event() {}
   
    handler_id operator+=(handler&& h)
    {
      std::lock_guard<std::mutex> lock(_lock);

      auto& val = _handlers.emplace_back(std::forward<handler>(h));
      return &val;
    }

    bool operator-=(handler_id id)
    {
      std::lock_guard<std::mutex> lock(_lock);

      for (auto h = _handlers.begin(); h != _handlers.end(); ++h)
      {
        if (&(*h) == id)
        {
          _handlers.erase(h);
          return true;
        }
      }
      return false;
    }

    auto bind(handler&& h)
    {
      return std::shared_ptr<const handler>(
        (*this) += std::forward<handler>(h), 
        [this](handler_id id) { (*this) -= id; });
    }

    R fire(Args&&... args) const
    {
      if (_handlers.empty())
        return R();

      _lock.lock();
      std::vector<handler> copy(_handlers.begin(), _handlers.end());
      _lock.unlock();

      return _collector(copy, std::forward<Args>(args)...);
    }

  private:
    std::list<handler> _handlers;
    mutable std::mutex _lock;
    TCollector _collector;
  };

  /// <summary>
  /// Event triggered when the xlOil addin is loaded by Excel
  /// </summary>
  XLOIL_EXPORT Event<void(void), VoidCollector>& 
    Event_AutoOpen();

  /// <summary>
  /// Event triggered when the xlOil addin is unloaded by Excel.
  /// Purposely not exported as plugins should unload when requested
  /// by xlOil Core, hence have not need to hook this event.
  /// </summary>
  Event<void(void), VoidCollector>& 
    Event_AutoClose();

  /// <summary>
  /// Event triggered at the end of an Excel calc cycle. Equivalent to
  /// Excel's Application.AfterCalculate event.
  /// </summary>
  XLOIL_EXPORT Event<void(void), VoidCollector>&
    Event_CalcEnded();

  /// <summary>
  /// Event triggered when calculation is cancelled by user interaction
  /// with Excel.
  /// </summary>
  XLOIL_EXPORT Event<void(void), VoidCollector>&
    Event_CalcCancelled();

  /// <summary>
  /// Event triggered when a workbook file is opened from storage. Passes
  /// the file path and file name as arguments. Equivalent to Excel's
  /// Application.WorkbookOpen event.
  /// </summary>
  XLOIL_EXPORT Event<void(const wchar_t*, const wchar_t*), VoidCollector>&
    Event_WorkbookOpen();

  /// <summary>
  /// Event triggered when a new workbook is created. Passes the
  /// workbook name as argument. Equivalent to Excel's
  /// Application.NewWorkbook event.
  /// </summary>
  XLOIL_EXPORT Event<void(const wchar_t*), VoidCollector>&
    Event_NewWorkbook();
  XLOIL_EXPORT Event<void(const wchar_t*), VoidCollector>&
    Event_WorkbookClose();
  
  enum class FileAction
  {
    /// Sent when a file is created or renamed
    Add = 1,
    /// Sent when a file is deleted or renamed
    Delete = 2,
    /// Sent when a file is modified
    Modified = 4
  };

  XLOIL_EXPORT Event<void(const wchar_t*, const wchar_t*, FileAction), VoidCollector>& 
    Event_DirectoryChange(const std::wstring& path);
}