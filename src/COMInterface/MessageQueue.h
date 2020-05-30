#pragma once
#include <functional>

namespace xloil
{
  namespace COM
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
    void queueWindowMessage(const std::function<void()>& func);
  }
}