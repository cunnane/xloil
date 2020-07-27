#pragma once
#include "ExportMacro.h"
#include <functional>
#include <future>
namespace xloil
{
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
      WINDOW  = 1 << 0,
      APC     = 1 << 1,
      ENQUEUE = 1 << 2,
      XLL_API = 1 << 3
    };
  }

  /// <summary>
  /// Excel will sometimes reject COM calls with the error VBA_E_IGNORE. This is
  /// against standard COM practice, but the COM interface, unlike the GUI icons
  /// has been abandoned and does not receive updates. xlOil with throw the
  /// ComBusyException in this case. Use of <see cref="excelApiCall"> catches
  /// this exception and retries, but with sufficent failures it make be passed
  /// to the user.
  /// </summary>
  class ComBusyException : public std::runtime_error
  {
  public:
    ComBusyException(const char* message)
      : std::runtime_error(message)
    {}
  };

  /// <summary>
  /// Excel's COM interface must be called on the main thread, since it's
  /// single-threaded apartment. This function queues an asynchronous procedure
  /// call (APC) callback or a message to a hidden window which will be executed
  /// on the main thread when Excel pumps its message loop or Excel passes control
  /// to windows respectively.  APC is generally executed when the user interacts 
  /// with Excel and window messages immediately after calculation.
  /// 
  /// Calls to the XLL interface require the main thread (with some exceptions) and
  /// being in the correct 'context', i.e. being in a function invoked by Excel.
  /// Setting the XLL_API flag schedules a callback to ensure this. 
  /// </summary>

  XLOIL_EXPORT std::future<void> 
    excelApiCall(
      const std::function<void()>& func, 
      int flags = QueueType::WINDOW, 
      int nRetries = 10, 
      unsigned waitTime = 200);
}