#include <xlOil/ExcelThread.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil-COM/XllContextInvoke.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComVariant.h>
#include <xloil/Log.h>
#include <xloil/AppObjects.h>
#include <xloil/Throw.h>
#include <xloil/State.h>
#include <xloil/ExcelUI.h>
#include <functional>
#include <queue>
#include <mutex>
#include <future>
#include <comdef.h>

using std::scoped_lock;
using std::shared_ptr;
using std::make_shared;
using std::vector;

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
      wc.hInstance     = (HINSTANCE)App::internals().hInstance;
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

    struct QueueItem
    {
      std::function<void()> _func;
      int _flags;
      unsigned _waitTime;

      QueueItem(
        const std::function<void()>& func, 
        int flags,
        unsigned waitTime)
        : _func(func)
        , _flags(flags)
        , _waitTime(waitTime)
      {}

      bool useCOM() const noexcept 
      {
        return (_flags & ExcelRunQueue::COM_API) != 0;
      }
      bool useXLL() const noexcept 
      {
        return (_flags & ExcelRunQueue::XLL_API) != 0;
      }
      bool operator()(bool comAvailable, bool xllAvailable) noexcept
      {
        if (useCOM() && !comAvailable)
          return false;
        if (useXLL() && !(xllAvailable || comAvailable))
          return false;
        // The _func should be a packaged task which is noexcept, so the only errors we catch
        // should come from runInXllContext.
        try
        {
          if (useXLL())
            runInXllContext(_func);
          else
            _func();
        }
        catch (const xloil::ComBusyException&)
        {
          // Even though we previously called the COM interface, it can still
          // become 'busy' later
          return false;
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Error running on main thread: {}", e.what());
        }
        catch (...)
        {
          XLO_ERROR("Error running on main thread: unknown");
        }
        return true;
      }
    };

    // Entirely arbitrary ID numbers
    static constexpr unsigned IDT_TIMER1     = 101;
    static constexpr unsigned WINDOW_MESSAGE = 666;
    static constexpr unsigned WM_TIMER       = 0x0113;

    auto firstJobTime(ULONGLONG now) 
    {
      // The queue is a sorted map so first element is due first.
      return _timerQueue.begin()->first > now
        ? unsigned(_timerQueue.begin()->first - now)
        : 0;
    }

    void startTimer(unsigned millisecs)
    {
      if (millisecs == 0)
        PostMessage(_hiddenWindow, WINDOW_MESSAGE, 0, 0);
      else
        SetTimer(_hiddenWindow, IDT_TIMER1, millisecs, TimerCallback);
    }

    void enqueue(const shared_ptr<QueueItem>& item, unsigned millisecs) noexcept
    {
      try
      {
        ULONGLONG now = GetTickCount64();
        {
          scoped_lock lock(_lock);
          _timerQueue.emplace(now + millisecs, item);
          if (millisecs > 0)
            millisecs = firstJobTime(now);
        }
        startTimer(millisecs);
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Internal error adding main thread queue item: '{}'", e.what());
      }
      catch (...)
      {
        XLO_ERROR("Internal error adding main thread queue item");
      }
    }

  private:
    static void CALLBACK TimerCallback(
      HWND /*hwnd*/, UINT /*uMsg*/, UINT_PTR /*idEvent*/, DWORD /*dwTime*/) noexcept
    {
      try
      {
        auto& self = instance();
        auto now = GetTickCount64();
        vector<shared_ptr<QueueItem>> items;
        {
          scoped_lock lock(self._lock);
          auto i = self._timerQueue.begin();
          auto end = self._timerQueue.end();
          // Find all the queue items with a due time before now and copy them
          // to our pending vector.
          while (i != end && i->first <= now) {
            items.push_back(i->second);
            ++i;
          }
          // Erase all the items copied to the pending vector
          self._timerQueue.erase(self._timerQueue.begin(), i);
        }

        // Nothing to do, then exit
        if (items.empty())
          return;

        // We have released mutex, now run pending queue items
        const auto comAvailable = COM::isComApiAvailable();
        const auto xllAvailable = InXllContext::check();

        items.erase(std::remove_if(items.begin(), items.end(),
          [=](auto& pJob)
          {
              return (*pJob)(comAvailable, xllAvailable);
          }), items.end());

        // Any remaining items failed due to COM/XLL availability
        // so are requeued.
        if (!items.empty())
        {
          now = GetTickCount64();
          unsigned startTime;
          {
            scoped_lock lock(self._lock);
            for (auto& item : items)
              if ((item->_flags & ExcelRunQueue::NO_RETRY) == 0)
                self._timerQueue.emplace(now + item->_waitTime, item);
            startTime = self.firstJobTime(now);
          }
          self.startTimer(startTime);
        }
      }
      catch (const std::exception& e)
      {
        XLO_ERROR("Internal error processing main thread queue: {}", e.what());
      }
      catch (...)
      {
        XLO_ERROR("Internal error running main thread queue: unknown");
      }
    }

    static LRESULT CALLBACK WindowProc(
      HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) noexcept
    {
      switch (uMsg)
      {
      case WINDOW_MESSAGE:
      case WM_TIMER:
      {
        TimerCallback(hwnd, uMsg, wParam, 0);
        return S_OK;
      }
      default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
      }
    }

    std::multimap<ULONGLONG, shared_ptr<QueueItem>> _timerQueue;
    std::mutex _lock;

    HWND _hiddenWindow;
    HANDLE _threadHandle;
  };

  void initMessageQueue()
  {
    Messenger::instance();
  }

  bool isMainThread()
  {
    return App::internals().mainThreadId == GetCurrentThreadId();
  }

  void runExcelThreadImpl(
    std::function<void()>&& func,
    int flags, 
    unsigned waitBeforeCall,
    unsigned waitBetweenRetries)
  {
    auto queueItem = make_shared<Messenger::QueueItem>(func,
      flags,
      waitBetweenRetries);

    // Try to run immediately if possible
    const bool canRunNow = waitBeforeCall == 0 
      && (flags & ExcelRunQueue::ENQUEUE) == 0 
      && isMainThread();
    if (canRunNow)
    {
      // TODO: avoid running isComApiAvailable if we don't need it? 
      // Generally functions scheduled for the main thread do need the COM or XLL interface
      const auto comAvailable = COM::isComApiAvailable();
      const auto xllAvailable = InXllContext::check();
      if ((*queueItem)(comAvailable, xllAvailable))
        return;
    }

    auto& messenger = Messenger::instance();
    messenger.enqueue(queueItem, waitBeforeCall);
  }

  struct RetryAtStartup
  {
    void operator()()
    {
      try
      {
        COM::connectCom();
        runExcelThread(func, ExcelRunQueue::XLL_API);
      }
      catch (const ComConnectException&)
      {
        XLO_DEBUG("Could not connect COM: trying again in 1 second...");
        runExcelThread(
          RetryAtStartup{ func },
          ExcelRunQueue::ENQUEUE | ExcelRunQueue::NO_RETRY,
          1000); // wait 1 second before call
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
    runExcelThread(RetryAtStartup{ func }, ExcelRunQueue::ENQUEUE);
  }
}