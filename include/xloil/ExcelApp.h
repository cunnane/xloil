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
  /// Internal use: called during Core DLL startup.
  /// </summary>
  void initMessageQueue();
  
  /// <summary>
  /// Returns true if the current thread is the main Excel thread
  /// </summary>
  XLOIL_EXPORT bool isMainThread();

  // Enum class just doesn't cut it with flags. You can overload all the 
  // operators and use some SFINAE to make it safe... or you can do this.
  namespace QueueType
  {
    enum QueueTypeValues
    {
      WINDOW  = 1 << 0, /// Run item via a hidden window message 
      APC     = 1 << 1, /// Run item via an APC call
      ENQUEUE = 1 << 2, /// Always queue item, do not try to run immediately
      XLL_API = 1 << 3  /// Item uses XLL API functions
    };
  }

  /// <summary>
  /// Excel will sometimes reject COM calls with the error VBA_E_IGNORE. This is
  /// against standard COM practice, but the COM interface, unlike the GUI,
  /// has been abandoned and does not receive updates. xlOil with throw the
  /// ComBusyException in this case. Use of <see cref="excelPost"> catches
  /// this exception and retries, but with sufficent failures it make be passed
  /// to the user.
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
  /// Excel's COM interface must be called on the main thread, since it's
  /// single-threaded apartment. This function queues an asynchronous procedure
  /// call (APC) callback or a message to a hidden window which will be executed
  /// on the main thread when Excel pumps its message loop or Excel passes control
  /// to windows respectively.  APC is generally executed when the user interacts 
  /// with Excel and window messages immediately after calculation. Generally the
  /// windows message is the most responsive choice.
  /// 
  /// Calls to the XLL interface require the main thread (with some exceptions) and
  /// being in the correct 'context', i.e. being in a function invoked by Excel.
  /// Setting the XLL_API flag schedules a callback to ensure this. 
  /// </summary>

  XLOIL_EXPORT std::future<void> 
    excelPost(
      const std::function<void()>& func, 
      int flags = QueueType::WINDOW, 
      int nRetries = 10, 
      unsigned waitBetweenRetries = 200,
      unsigned waitBeforeCall = 0);

  
  void xllOpenComCall(const std::function<void()>& func);
}