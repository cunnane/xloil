#pragma once
#include "ExportMacro.h"
#include <xlOil/Log.h> // TODO: slim down this header!
#include <functional>
#include <memory>
#include <list>
#include <mutex>
#include <future>
#include <string>
#include <boost/preprocessor/stringize.hpp>

namespace xloil { class Range; }

namespace xloil
{
  namespace detail
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
        {
          try
          {
            h(args...);
          }
          catch (const std::exception& e)
          {
            XLO_ERROR("Error during event: {}", e.what());
          }
        }
      }
    };
  }
  namespace Event
  {
    template<class, class = detail::VoidCollector> class Event {};

    /// <summary>
    /// An observer-pattern based Event handler
    /// </summary>
    template<class R, class TCollector, class... Args>
    class Event<R(Args...), TCollector> :
      public std::enable_shared_from_this<Event<R(Args...), TCollector>>
    {
    public:
      using handler = std::function<R(Args...)>;
      using handler_id = const handler*;

      Event(const char* name = 0)
        : _name(name ? name : "?")
      {}

      /// <summary>
      /// Registers an event handler
      /// </summary>
      /// <param name="h"></param>
      /// <returns>An ID which can be used to unregister the handler</returns>
      handler_id operator+=(handler&& h)
      {
        std::lock_guard<std::mutex> lock(_lock);

        auto& val = _handlers.emplace_back(std::forward<handler>(h));
        return &val;
      }

      /// <summary>
      /// Removes an event handler give its registration ID.
      /// </summary>
      /// <param name="id"></param>
      /// <returns></returns>
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

      /// <summary>
      /// Removes an event handler given a reference to the handler
      /// </summary>
      bool operator-=(const handler& id)
      {
        return (*this) -= &id;
      }

      /// <summary>
      /// Registers an event handler and returns a shared_ptr whose destructor
      /// unregisters the event handler. The dtor keeps a shared_ptr to this 
      /// event to ensure the correct order of static destruction.
      /// </summary>
      /// <param name="h"></param>
      /// <returns></returns>
      auto bind(handler&& h)
      {
        auto thisPtr = shared_from_this();
        return std::shared_ptr<const handler>(
          (*this) += std::forward<handler>(h),
          [thisPtr](handler_id id) { (*thisPtr) -= id; });
      }

      R fire(Args&&... args) const
      {
        if (_handlers.empty())
          return R();

        _lock.lock();
        std::vector<handler> copy(_handlers.begin(), _handlers.end());
        _lock.unlock();
        XLO_TRACE("Firing event {0}", _name);
        return _collector(copy, std::forward<Args>(args)...);
      }

      const std::list<handler>& handlers() const 
      {
        return _handlers;
      }

      /// <summary>
      /// Removes all existing handlers
      /// </summary>
      void clear()
      {
        _handlers.clear();
      }

      const std::string& name() const { return _name; }

    private:
      std::list<handler> _handlers;
      mutable std::mutex _lock;
      TCollector _collector;
      std::string _name;
    };

    using EventXll = Event<void(void), detail::VoidCollector>;
    using EventNameParam = Event<void(const wchar_t*), detail::VoidCollector>;

    /// <summary>
    /// Event triggered when the xlOil addin is unloaded by Excel.
    /// Purposely not exported as plugins should unload when requested
    /// by xlOil Core, hence have not need to hook this event.
    /// </summary>
    EventXll&
      AutoClose();

    /// <summary>
    /// Event triggered at the end of an Excel calc cycle. Equivalent to
    /// Excel's Application.AfterCalculate event.
    /// </summary>
    XLOIL_EXPORT EventXll&
      AfterCalculate();

    /// <summary>
    /// Event triggered when calculation is cancelled by user interaction
    /// with Excel.
    /// </summary>
    XLOIL_EXPORT EventXll&
      CalcCancelled();

    /// <summary>
    /// Triggered when a new workbook is created. Passes the
    /// workbook name as argument.  See the Excel
    /// Application.NewWorkbook event.
    /// </summary>
    XLOIL_EXPORT EventNameParam&
      NewWorkbook();

    XLOIL_EXPORT Event<void(const wchar_t* wsName, const Range& target)>&
      SheetSelectionChange();

    XLOIL_EXPORT Event<void(const wchar_t* wsName, const Range& target, bool& cancel)>&
      SheetBeforeDoubleClick();

    XLOIL_EXPORT Event<void(const wchar_t* wsName, const Range& target, bool& cancel)>&
      SheetBeforeRightClick();

    XLOIL_EXPORT EventNameParam&
      SheetActivate();

    XLOIL_EXPORT EventNameParam&
      SheetDeactivate();

    XLOIL_EXPORT EventNameParam&
      SheetCalculate();

    XLOIL_EXPORT Event<void(const wchar_t* wsName, const Range& target)>&
      SheetChange();

    /// <summary>
    /// Triggered when a workbook file is opened from storage. Passes
    /// the file path and file name as arguments. See the Excel
    /// Application.WorkbookOpen event.
    /// </summary>
    XLOIL_EXPORT Event<void(const wchar_t* wbPath, const wchar_t* wbName)>&
      WorkbookOpen();

    /// <summary>
    /// WorkbookAfterClose is a special event: Excel's event `WorkbookBeforeClose`, is 
    /// limited by being cancellable by the user: it is not possible to know if the 
    /// workbook actually closed. When xlOil calls `WorkbookAfterClose`, the workbook is
    /// certainly closed, but it may be some time since that closure happened.
    /// 
    /// The event is not called for each workbook when xlOil exits.
    /// </summary>
    XLOIL_EXPORT EventNameParam&
      WorkbookAfterClose();

    XLOIL_EXPORT EventNameParam&
      WorkbookActivate();

    XLOIL_EXPORT EventNameParam&
      WorkbookDeactivate();

    XLOIL_EXPORT Event<void(const wchar_t* wbName, bool& cancel)>&
      WorkbookBeforeClose();

    XLOIL_EXPORT Event<void(const wchar_t* wbName, bool saveAsUI, bool& cancel)>&
      WorkbookBeforeSave();

    XLOIL_EXPORT Event<void(const wchar_t* wbName, bool& cancel)>&
      WorkbookBeforePrint();

    XLOIL_EXPORT Event<void(const wchar_t* wbName, const wchar_t* wsName)>&
      WorkbookNewSheet();

    XLOIL_EXPORT EventNameParam&
      WorkbookAddinInstall();

    XLOIL_EXPORT EventNameParam&
      WorkbookAddinUninstall();

    enum class FileAction
    {
      /// Sent when a file is created or renamed
      Add = 1,
      /// Sent when a file is deleted or renamed
      Delete = 2,
      /// Sent when a file is modified
      Modified = 4
    };

    XLOIL_EXPORT Event<
      void(const wchar_t* directory, const wchar_t* filename, FileAction)> &
      DirectoryChange(const std::wstring& path);

    XLOIL_EXPORT void allowEvents(bool value);

  }

 

  /// <summary>
  /// All the singleton xlOil events as a sequence. Use BOOST_PP_SEQ functions
  /// to iterate over this sequence to create bindings. Non singleton/static
  /// events such as DirectoryChange as not included here.
  /// </summary>
#define XLOIL_STATIC_EVENTS \
    (AfterCalculate)\
    (WorkbookOpen)\
    (NewWorkbook)\
    (SheetSelectionChange)\
    (SheetBeforeDoubleClick)\
    (SheetBeforeRightClick)\
    (SheetActivate)\
    (SheetDeactivate)\
    (SheetCalculate)\
    (SheetChange)\
    (WorkbookAfterClose)\
    (WorkbookActivate)\
    (WorkbookDeactivate)\
    (WorkbookBeforeClose)\
    (WorkbookBeforeSave)\
    (WorkbookBeforePrint)\
    (WorkbookNewSheet)\
    (WorkbookAddinInstall)\
    (WorkbookAddinUninstall)


    /// <summary>
    /// Creates a function body for a event declaration such as
    /// <code>
    ///   Event<void(void)>& MyEvent();
    /// </code>
    /// You need to include BOOST_PP_STRINGIZE to use this macro.
    /// </summary>
#define XLOIL_DEFINE_EVENT(name) \
    decltype(name()) name() \
    { \
      static auto e = std::make_shared<std::remove_reference_t<\
        decltype(name())>>(BOOST_PP_STRINGIZE(name)); \
      return *e; \
    };
}