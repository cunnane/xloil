#pragma once
#include "ExportMacro.h"
#include <functional>
#include <future>
namespace xloil
{
  void initMessageQueue();
  /// <summary>
  /// Excel's COM interface must be called on the main thread, since it's
  /// single-threaded apartment. This function queues an asynchronous 
  /// procedure call (APC) callback which will be executed on the main thread 
  /// when Excel pumps its message loop (typically when the user interacts 
  /// with Excel).
  /// </summary>
  void queueAPC(const std::function<void()>& func);

  /// <summary>
  /// Excel's COM interface must be called on the main thread, since it's
  /// single-threaded apartment. This function queues a callback via a 
  /// message to a hidden window. This will be executed on the main thread 
  /// when Excel's passes control to Windows (typically immediately after 
  /// calculation).
  /// </summary>
  XLOIL_EXPORT void queueWindowMessage(const std::function<void()>& func);
    
  XLOIL_EXPORT bool isMainThread();

  template <class TReturn>
  std::future<TReturn> runOnMainThread(const std::function<TReturn()>& func)
  {
    auto promise = std::make_shared<std::promise<TReturn>>();
    if (isMainThread())
      promise->set_value(func());
    else
    {
      queueWindowMessage([promise, func]()
      {
        promise->set_value(func());
      });
    }
    return promise->get_future();
  }
}