#pragma once
#include <xlOil/ExcelApp.h>
#include <xlOil/WindowsSlim.h>
#include <xloil/Loaders/EntryPoint.h>
#include <xlOil-COM/XllContextInvoke.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComAddin.h>
#include <xlOil-COM/ComVariant.h>
#include <xloil/Log.h>
#include <xloil/Throw.h>
#include <xloil/State.h>
#include <xloil/Ribbon.h>
#include <functional>
#include <queue>
#include <mutex>
#include <future>
#include <comdef.h>

using std::scoped_lock;
using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  Excel::_Application& excelApp() noexcept
  {
    return COM::excelApp();
  }

  std::shared_ptr<IComAddin> xloil::makeComAddin(
    const wchar_t * name, const wchar_t * description)
  {
    return COM::createComAddin(name, description);
  }

  ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange)
  {
    return COM::variantToExcelObj(variant, allowRange);
  }
  void excelObjToVariant(VARIANT* v, const ExcelObj& obj)
  {
    COM::excelObjToVariant(v, obj);
  }
  class Messenger
  {
  public:
    Messenger()
    {
      _threadHandle = OpenThread(THREAD_SET_CONTEXT, true, GetCurrentThreadId());

      WNDCLASS wc;
      memset(&wc, 0, sizeof(WNDCLASS));
      wc.lpfnWndProc   = WindowProc;
      wc.hInstance     = (HINSTANCE)State::excelState().hInstance;
      wc.lpszClassName = L"xlOilHidden";
      if (RegisterClass(&wc) == 0)
        XLO_ERROR(L"Failed to register window class: {0}", writeWindowsError());

      _hiddenWindow = CreateWindow(
        wc.lpszClassName,
        L"",         // Window text
        0,           // Window style
        // Size and position (4 args)
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        HWND_MESSAGE, NULL,  // Parent window, Menu 
        wc.hInstance,
        NULL);

      if (!_hiddenWindow)
        XLO_ERROR(L"Failed to create window: {0}", writeWindowsError());
    }

    static Messenger& instance()
    {
      static Messenger obj;
      return obj;
    }

    struct QueueItem : std::enable_shared_from_this<QueueItem>
    {
      std::function<void()> _func;
      std::shared_ptr<std::promise<void>> _promise;
      int _flags;
      int _nComRetries;
      unsigned _waitTime;
      

      QueueItem(
        const std::function<void()>& func, 
        std::shared_ptr<std::promise<void>> promise,
        int flags,
        int nComRetries, 
        unsigned waitTime)
        : _func(func)
        , _promise(promise)
        , _flags(flags)
        , _nComRetries(nComRetries)
        , _waitTime(waitTime)
      {}

      void operator()(Messenger& messenger);
    };

    static constexpr unsigned WINDOW_MESSAGE = 666;
   
    void QueueAPC(const shared_ptr<QueueItem>& item)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _apcQueue.empty();
      _apcQueue.emplace_back(item);
      if (emptyQueue)
        QueueUserAPC(processAPCQueue, _threadHandle, (ULONG_PTR)this);
    }

    void QueueWindow(const shared_ptr<QueueItem>& item)
    {
      scoped_lock lock(_lock);
      const bool emptyQueue = _windowQueue.empty();
      _windowQueue.emplace_back(item);
      if (emptyQueue)
        PostMessage(_hiddenWindow, WINDOW_MESSAGE, (WPARAM)this, 0);
    }

    void queueWindowTimer(const shared_ptr<QueueItem>& item, int millisecs) noexcept
    {
      scoped_lock lock(_lock);
      _timerQueue[item.get()] = item; // is this noexcept?
      SetTimer(_hiddenWindow, (UINT_PTR)item.get(), millisecs, TimerCallback);
    }

  private:
    static void CALLBACK TimerCallback(
      HWND hwnd, UINT /*uMsg*/, UINT_PTR idEvent, DWORD /*dwTime*/)
    {
      auto& self = instance();
      auto retKill = KillTimer(hwnd, idEvent);
      shared_ptr<QueueItem> item;
      {
        scoped_lock lock(self._lock);
        auto found = self._timerQueue.find((QueueItem*)idEvent);
        if (found == self._timerQueue.end())
        {
          XLO_ERROR("Internal error: bad window timer");
          return;
        }
        item = found->second;
        self._timerQueue.erase(found);
      }
      (*item)(self);
    }

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
        (*job)(self);
      }
    }

    std::deque<shared_ptr<QueueItem>> _windowQueue;
    std::deque<shared_ptr<QueueItem>> _apcQueue;
    std::unordered_map<QueueItem*, shared_ptr<QueueItem>> _timerQueue;
    std::mutex _lock;

    HWND _hiddenWindow;
    HANDLE _threadHandle;
  };

  void initMessageQueue()
  {
    Messenger::instance();
  }

  void Messenger::QueueItem::operator()(Messenger& messenger)
  {
    if (_nComRetries > 0 && (_flags & ExcelRunQueue::COM_API) != 0 && !COM::isComApiAvailable())
    {
      try
      {
        --_nComRetries;
        //TODO: if _isAPC, then use SetWaitableTimer
        messenger.queueWindowTimer(shared_from_this(), _waitTime);
      }
      catch (...)
      {
        _promise->set_exception(make_exception_ptr(std::runtime_error("Unknown exception")));
      }
      return;
    }
   
    try
    {
      if ((_flags & ExcelRunQueue::XLL_API) != 0)
        runInXllContext(_func);
      else
        _func();
    }
    catch (const std::exception& e) // What about SEH?
    {
      _promise->set_exception(make_exception_ptr(e));
    }
    catch (_com_error& error)
    { 
      _promise->set_exception(
        make_exception_ptr(std::runtime_error(fmt::format(
          "COM Error {0:#x}", (size_t)error.Error()
        ))));
    }
    catch (...)
    {
      _promise->set_exception(make_exception_ptr(std::runtime_error("Unknown exception")));
    }
  }


  bool isMainThread()
  {
    // TODO: would a thread-local bool be quicker here?
    return State::excelState().mainThreadId == GetCurrentThreadId();
  }

  std::future<void> excelRunOnMainThread(
    const std::function<void()>& func, 
    int flags, 
    int nRetries, 
    unsigned waitBetweenRetries,
    unsigned waitBeforeCall)
  {
    auto promise = std::make_shared<std::promise<void>>();

    auto queueItem = make_shared<Messenger::QueueItem>([promise, func]()
      {
          func();
          promise->set_value();
      },
      promise, 
      flags,
      nRetries,
      waitBetweenRetries);

    auto& messenger = Messenger::instance();

    // Try to run immediately if possible
    if (waitBeforeCall == 0 && (flags & ExcelRunQueue::ENQUEUE) == 0 && isMainThread())
      (*queueItem)(messenger);
    else
    {
      // Otherwise for XLL API usage we also need the COM API to switch to XLL context
      if ((flags & ExcelRunQueue::XLL_API) != 0)
        queueItem->_flags |= ExcelRunQueue::COM_API;

      if (waitBeforeCall > 0)
        messenger.queueWindowTimer(queueItem, waitBeforeCall);
      else if ((flags & ExcelRunQueue::APC) != 0)
        messenger.QueueAPC(queueItem);
      else
        messenger.QueueWindow(queueItem);
    }
    return promise->get_future();
  }

  struct RetryAtStartup
  {
    void operator()()
    {
      try
      {
        COM::connectCom();
        excelRunOnMainThread(func, ExcelRunQueue::XLL_API);
      }
      catch (const COM::ComConnectException&)
      {
        excelRunOnMainThread(
          RetryAtStartup{ func },
          ExcelRunQueue::WINDOW | ExcelRunQueue::ENQUEUE,
          0, // no retry
          0,
          1000 // wait 1 second before call
        );
      }
      catch (const std::exception& e)
      {
        XLO_ERROR(e.what());
      }
    }
    std::function<void()> func;
  };

  void runComSetupOnXllOpen(const std::function<void()>& func)
  {
    excelRunOnMainThread(RetryAtStartup{ func }, ExcelRunQueue::ENQUEUE);
  }
}