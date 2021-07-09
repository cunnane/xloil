#pragma once
#include "ExportMacro.h"
#include <functional>
#include <future>

namespace Excel { struct _Application; }


namespace xloil
{
  /// <summary>
  /// Gets the Excel.Application object which is the root of the COM API 
  /// </summary>
  XLOIL_EXPORT Excel::_Application& excelApp() noexcept;

  /// <summary>
  /// Returns true if the current thread is the main Excel thread
  /// </summary>
  XLOIL_EXPORT bool isMainThread();

  /// <summary>
  /// Determines how <see cref="excelRunOnMainThread"/> will dispatch the provided
  /// function. 
  /// </summary>
  namespace ExcelRunQueue
  {
    enum QueueTypeValues
    {
      /// Run item via a hidden window message (default)
      WINDOW  = 1 << 0,
      /// Run item via an APC call (use in special cases)
      APC     = 1 << 1, 
      /// Always queue item, do not try to run immediately if on main thread
      ENQUEUE = 1 << 2, 
      /// Item uses XLL API functions
      XLL_API = 1 << 3, 
      /// Item uses COM API functions (default)
      COM_API = 1 << 4  
    };
  }

  /// <summary>
  /// Excel will sometimes reject COM calls with the error VBA_E_IGNORE. This can 
  /// happen when the user needs to complete a UI gesture such as closing dialog box.
  /// This is against standard COM practice, but the COM interface, unlike the GUI colours
  /// has been abandoned and does not appear to receive updates. xlOil will throw the
  /// ComBusyException in this case. Use of <see cref="excelRunOnMainThread"/> allows
  /// retrying a call until the interface becomes available.
  /// </summary>
  class ComBusyException : public std::runtime_error
  {
  public:
    ComBusyException(const char* message = nullptr)
      : std::runtime_error(message 
        ? message 
        : "Excel COM is busy; a dialog box may be open. Retry the action and if this error persists, restart Excel.")
    {}
  };


  /// <summary>
  /// Excel's COM interface, that is any called based on the Excel::Application object,
  /// must be called on the main thread. This function schedules a callback from the
  /// main thread or executes immediately if called from the main thread (a  
  /// callback can be forced with the ENQUEUE flag).
  /// 
  /// The COM API is not always available - see <see cref="ComBusyException"/> for
  /// a discussion of this issue.
  /// 
  /// Calls to the XLL interface require the main thread (with some exceptions) and
  /// being in the correct 'context', i.e. being in a function invoked by Excel.
  /// Setting the ExcelRunQueue::XLL_API flag schedules a callback to ensure this. 
  /// xlOil uses the COM interface to switch to XLL context, so consider setting 
  /// non-zero COM retries if using this option.
  /// 
  /// </summary>
  /// <param name="func"></param>
  /// <param name="flags">
  ///   The callback can be schedule via an asynchronous procedure call (APC)
  ///   or a message to a hidden window which will be executed.  APC is generally 
  ///   executed when the user interacts with Excel and window messages immediately 
  ///   after calculation. Generally the Windows message (the default) is the more 
  ///   responsive choice.
  /// </param>
  /// <param name="nComRetries">
  ///   If non zero, xlOil will check if the COM API is available before
  ///   invoking `func`. If unavailable, the retry count is decremented and xlOil
  ///   tries again in <param ref="waitBetweenRetries"/> milliseconds.  If you do
  ///   not intend to call any COM functions, this can be set to zero.
  /// </param>
  /// <param name="waitBetweenRetries">
  ///   The number of milliseconds to wait between COM retries if Excel reports
  ///   that the COM API is not available.
  /// </param>
  /// <param name="waitBeforeCall">
  ///   Wait for the specified number of milliseconds before executing the callback
  /// </param>
  /// <returns>A std::future which contains the result of <param ref="func"/></returns>
  /// 
  XLOIL_EXPORT std::future<void> 
    excelRunOnMainThread(
      const std::function<void()>& func, 
      int flags = ExcelRunQueue::WINDOW | ExcelRunQueue::COM_API,
      int nComRetries = 10, 
      unsigned waitBetweenRetries = 200,
      unsigned waitBeforeCall = 0);

  
  void runComSetupOnXllOpen(const std::function<void()>& func);

  /// <summary>
  /// Internal use: called during Core DLL startup.
  /// </summary>
  void initMessageQueue();
}