#pragma once
#include <xloil/ApiMessage.h>
#include <xlOil/WindowsSlim.h>
#include <xloil/Loaders/EntryPoint.h>
#include <COMInterface/XllContextInvoke.h>
#include <xloil/Log.h>
#include <xloil/State.h>
#include <functional>
#include <queue>
#include <mutex>
#include <future>

using std::scoped_lock;
using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  class Messenger
  {
  public:
    Messenger()
    {
      _threadHandle = OpenThread(THREAD_SET_CONTEXT, true, GetCurrentThreadId());

      WNDCLASS wc;
      memset(&wc, 0, sizeof(WNDCLASS));
      wc.lpfnWndProc   = WindowProc;
      wc.hInstance     = (HINSTANCE)State::excelHInstance();
      wc.lpszClassName = L"xlOilHidden";
      if (RegisterClass(&wc) == 0)
        XLO_ERROR(L"Failed to register window class: {0}", writeWindowsError());

      _hiddenWindow = CreateWindow(
        wc.lpszClassName,
        L"",                 // Window text
        0,                   // Window style
        // Size and position
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        NULL, NULL,          // Parent window, Menu 
        wc.hInstance,
        NULL);

      if (!_hiddenWindow)
        XLO_ERROR(L"Failed to create window: {0}", writeWindowsError());
    }

    struct QueueItem : std::enable_shared_from_this<QueueItem>
    {
      std::function<void()> _func;
      std::shared_ptr<std::promise<void>> _promise;
      bool _usesXllApi;
      int _nRetries; 
      unsigned _waitTime;
      bool _isAPC;

      QueueItem(const std::function<void()>& func, std::shared_ptr<std::promise<void>> promise, bool usesXllApi, int nRetries, unsigned waitTime, bool isAPC)
        : _func(func)
        , _promise(promise)
        , _usesXllApi(usesXllApi)
        , _nRetries(nRetries)
        , _waitTime(waitTime)
        , _isAPC(isAPC)
      {}

      void operator()();
    };

    static constexpr unsigned WINDOW_MESSAGE = 666;

    void QueueAPC(shared_ptr<QueueItem> item)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _apcQueue.empty();
      _apcQueue.emplace_back(item);
      if (emptyQueue)
        QueueUserAPC(processAPCQueue, _threadHandle, (ULONG_PTR)this);
    }

    void QueueWindow(shared_ptr<QueueItem> item)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _windowQueue.empty();
      _windowQueue.emplace_back(item);
      if (emptyQueue)
        PostMessage(_hiddenWindow, WINDOW_MESSAGE, (WPARAM)this, 0);
    }

  private:
    static LRESULT CALLBACK WindowProc(
      HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
      switch (uMsg)
      {
      case WINDOW_MESSAGE:
      {
        processWindowQueue((ULONG_PTR)wParam);
        return S_OK;
      }
      default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
      }
    }

    static void processWindowQueue(ULONG_PTR ptr)
    {
      auto& self = *(Messenger*)ptr;
      processQueue(self, self._windowQueue);
    }

    static void __stdcall processAPCQueue(ULONG_PTR ptr)
    {
      auto& self = *(Messenger*)ptr;
      processQueue(self, self._apcQueue);
    }

    static void processQueue(Messenger& self, std::deque<shared_ptr<QueueItem>>& queue)
    {
      decltype(_apcQueue) jobs;
      {
        scoped_lock lock(self._lock);
        jobs.assign(queue.begin(), queue.end());
        queue.clear();
      }

      for (auto& job : jobs)
      {
        (*job)();
      }
    }

    std::deque<shared_ptr<QueueItem>> _windowQueue;
    std::deque<shared_ptr<QueueItem>> _apcQueue;
    std::mutex _lock;

    HWND _hiddenWindow;
    HANDLE _threadHandle;
  };

  Messenger& getMessenger()
  {
    static Messenger obj;
    return obj;
  }
  void initMessageQueue()
  {
    getMessenger();
  }

  void enqueue(shared_ptr<Messenger::QueueItem> item)
  {
    if (item->_isAPC)
      getMessenger().QueueAPC(item);
    else
      getMessenger().QueueWindow(item);
  }

  void Messenger::QueueItem::operator()()
  {
    try
    {
      if (_usesXllApi)
        runInXllContext(_func);
      else
        _func();
    }
    catch (const ComBusyException& e)
    {
      if (0 == _nRetries--)
        _promise->set_exception(make_exception_ptr(e));
      Sleep(_waitTime);
      enqueue(shared_from_this());
    }
    catch (const std::exception& e) // what about SEH?
    {
      _promise->set_exception(make_exception_ptr(e));
    }
  }


  bool isMainThread()
  {
    // TODO: would a thread-local bool be quicker here?
    return State::mainThreadId() == GetCurrentThreadId();
  }

  std::future<void> excelApiCall(const std::function<void()>& func, int flags, int nRetries, unsigned waitTime)
  {
    auto promise = std::make_shared<std::promise<void>>();

    auto queueItem = make_shared<Messenger::QueueItem>([promise, func]()
      {
          func();
          promise->set_value();
      },
      promise, (flags & (int)QueueType::XLL_API), nRetries, waitTime, (flags & (int)QueueType::APC));

    if ((flags & (int)QueueType::ENQUEUE) == 0 && isMainThread())
      (*queueItem)();
    else
      enqueue(queueItem);

    return promise->get_future();
  }
}