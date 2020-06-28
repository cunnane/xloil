#pragma once
#include <xloil/ThreadControl.h>
#include <xlOilHelpers/WindowsSlim.h>
#include <xloil/Loaders/EntryPoint.h>
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
    using QueueItem = std::function<void()>;

    static constexpr unsigned WINDOW_MESSAGE = 666;

    void QueueAPC(const QueueItem& func)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _queue.empty();
      _queue.emplace_back(new QueueItem(func));
      if (emptyQueue)
        QueueUserAPC(processQueue, _threadHandle, (ULONG_PTR)this);
    }

    void QueueWindow(const QueueItem& func)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _queue.empty();
      _queue.emplace_back(new QueueItem(func));
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
        processQueue((ULONG_PTR)wParam);
        return S_OK;
      }
      default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
      }
    }

    static void processQueue(ULONG_PTR ptr)
    {
      auto& self = *(Messenger*)ptr;
      std::vector<shared_ptr<QueueItem>> jobs;
      {
        scoped_lock lock(self._lock);
        jobs.assign(self._queue.begin(), self._queue.end());
        self._queue.clear();
      }

      for (auto& job : jobs)
        (*job)();
    }

    std::deque<shared_ptr<QueueItem>> _queue;
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
  void queueAPC(const std::function<void()>& func)
  {
    getMessenger().QueueAPC(func);
  }
  void queueWindowMessage(const std::function<void()>& func)
  {
    getMessenger().QueueWindow(func);
  }
  bool isMainThread()
  {
    return State::mainThreadId() == GetCurrentThreadId();
  }
}