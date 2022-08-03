#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include <xloil/LogWindow.h>
#include <xloil/ExcelThread.h>
#include <xloil/StringUtils.h>
#include <xlOilHelpers/Environment.h>
#include <xlOilHelpers/Exception.h>
#include <list>

using std::wstring;
using std::string;
using std::shared_ptr;
using std::make_shared;

namespace xloil
{
  class LogWindow : public ILogWindow
  {
    HWND theMainWindow;
    HWND theTextControl;
    HMENU theTextControlId = (HMENU)101;
    WNDPROC theMenuHandler;
    bool theWindowIsOpen = false;
    std::list<wstring> _messages;
    wstring _windowText;
    size_t _maxSize;
    shared_ptr<ATOM> _windowClass;
    static constexpr const wchar_t* theWindowClass = L"xlOil_Log";

    static auto createWindowClass(HINSTANCE hInstance, const wchar_t* windowClass)
    {
      // Define the main window class
      WNDCLASSEX win;
      win.cbSize = sizeof(WNDCLASSEX);
      win.hInstance = hInstance;
      win.lpszClassName = windowClass;
      win.lpfnWndProc = StaticWindowProc;
      win.style = CS_HREDRAW | CS_VREDRAW;

      // Use default icons and mouse pointer
      win.hIcon = LoadIcon(NULL, IDI_APPLICATION);
      win.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
      win.hCursor = LoadCursor(NULL, IDC_ARROW);

      win.lpszMenuName = NULL;
      win.cbClsExtra = 0;  // No extra bytes after the window class
      win.cbWndExtra = sizeof(void*);
      win.hbrBackground = GetSysColorBrush(COLOR_3DFACE); // Use default colour

      auto atom = RegisterClassEx(&win);

      return std::shared_ptr<ATOM>(new ATOM(atom), [](ATOM* atom) {
        UnregisterClass((LPCWSTR)LOWORD(*atom), nullptr);
        delete atom;
      });
    }

  public:
    LogWindow(
      HWND parentWnd,
      HINSTANCE hInstance,
      const wchar_t* winTitle,
      HMENU menuBar,
      WNDPROC menuHandler,
      size_t maxSize)
      : _maxSize(maxSize)
    {
      static auto winClass = createWindowClass(hInstance, theWindowClass);

      // Take a reference to the window class so we never try to unregister the
      // class while there are open windows remaining.
      _windowClass = winClass;

      // If we try to create a window from our class during xlAutoOpen, we will get
      // the cryptic "ntdll.dll (EXCEL.EXE) RangeChecks instrumentation code detected 
      // an out of range array access".  Whatever Excel gets up to during start-up
      // seems to screw around with the Win32 API.  We use the existence or 
      // otherwise of a parentHwnd to determine if we are in this perilous state
      auto hwnd = FindWindow(theWindowClass, winTitle);
      if (!hwnd)
        hwnd = CreateWindowEx(
          0,
          parentWnd ? theWindowClass : L"EDIT",
          winTitle,
          parentWnd
          ? WS_OVERLAPPEDWINDOW  // Title bar, minimimise, close and resize controls
          : WS_OVERLAPPEDWINDOW | WS_VSCROLL | ES_LEFT | ES_MULTILINE,
          CW_USEDEFAULT, CW_USEDEFAULT, // (x, y)-position
          CW_USEDEFAULT, CW_USEDEFAULT, // width, height
          HWND_DESKTOP,
          menuBar,
          hInstance,
          this);

      if (!hwnd)
        throw Helpers::Exception(L"Failed to create LogWindow: %s",
          Helpers::writeWindowsError().c_str());

      if (!parentWnd)
        theTextControl = hwnd;

      theMenuHandler = menuHandler;
      theMainWindow = hwnd;
    }

  private:
    static LRESULT CALLBACK StaticWindowProc(
      HWND hwnd,
      UINT message,
      WPARAM wParam,
      LPARAM lParam)
    {
      LogWindow* instance;
      if (message == WM_CREATE)
      {
        auto create = (CREATESTRUCT*)lParam;
        instance = (LogWindow*)create->lpCreateParams;
        SetWindowLongPtr(hwnd, 0, (LONG_PTR)instance);
      }
      else
        instance = (LogWindow*)GetWindowLongPtr(hwnd, 0);

      if (instance)
        return instance->WindowProc(hwnd, message, wParam, lParam);

      return DefWindowProc(hwnd, message, wParam, lParam);
    }

    LRESULT CALLBACK WindowProc(
      HWND hwnd,
      UINT message,
      WPARAM wParam,
      LPARAM lParam)
    {
      constexpr int xOffset = 5, yOffset = 5;

      switch (message)
      {
      case WM_CREATE:
        theTextControl = CreateWindow(L"EDIT", // an edit control
          NULL,        // no window title 
          WS_CHILD | WS_VISIBLE | WS_VSCROLL |
          ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL,
          0, 0, 0, 0,  // we will set size in WM_SIZE message 
          hwnd,        // parent window 
          theTextControlId,
          (HINSTANCE)GetWindowLongPtr(hwnd, GWLP_HINSTANCE),
          NULL);       // pointer not needed 

        if (!theTextControl)
          return EXIT_FAILURE;

        return 0;

      case WM_SETFOCUS:
        SetFocus(theTextControl);
        return 0;

      case WM_SIZE:
        // Make the edit control the size of the window's client area. 
        MoveWindow(theTextControl,
          xOffset, yOffset,      // starting x- and y-coordinates 
          LOWORD(lParam) - xOffset,        // width of client area 
          HIWORD(lParam) - yOffset,        // height of client area 
          TRUE);                 // repaint window 
        return 0;

      case WM_CLOSE:
        LogWindow::theWindowIsOpen = false;
        ShowWindow(hwnd, SW_HIDE);
        return 0;

      case WM_COMMAND:
        if (theMenuHandler)
          return theMenuHandler(hwnd, message, wParam, lParam);
      }
      return DefWindowProc(hwnd, message, wParam, lParam);
    }

    // Just scroll to the end, word-wrap seems to confuse the
    // line count, so we just specify a big number for numLines
    void setTextBoxContents(const wchar_t* text, size_t numLines = 66666666)
    {
      // Add text to the window. 
      SendMessage(theTextControl, WM_SETTEXT, 0, (LPARAM)text);
      // Scroll to the last line
      SendMessage(theTextControl, EM_LINESCROLL, 0, numLines);
    }

    bool isOpen() const
    {
      return theWindowIsOpen;
    }

    void appendToWindow(const wstring& msg)
    {
      auto wmsg(msg);

      // Fix any unix line endings (e.g. from python)
      auto pos = (size_t)-1;
      while ((pos = wmsg.find(L'\n', pos + 1)) != string::npos)
      {
        if (wmsg[pos - 1] != L'\r')
          wmsg.replace(pos, 1, L"\r\n");
      }

      _windowText.append(wmsg).append(L"\r\n");
    }

  protected:
    virtual void showWindow() noexcept
    {
      setTextBoxContents(_windowText.c_str());
      ShowWindow(theMainWindow, SW_SHOWNORMAL);
      BringWindowToTop(theMainWindow);
      theWindowIsOpen = true;
    }

    virtual void setWindowText() noexcept
    {
      setTextBoxContents(_windowText.c_str());
    }

  public:
    void openWindow() noexcept override
    {
      if (isOpen())
        return;

      _windowText.clear();
      for (auto& msg : _messages)
        appendToWindow(msg);

      showWindow();
    }

    void appendMessage(wstring&& msg) noexcept override
    {
      _messages.emplace_back(std::forward<wstring>(msg));

      if (_messages.size() > _maxSize)
        _messages.pop_front();

      if (isOpen())
      {
        appendToWindow(_messages.back());
        setWindowText();
      }
    }
  };

  // Should only call window drawing API functions on the main thread
  class LogWindowThreaded : public LogWindow
  {
  public:
    using LogWindow::LogWindow;

  protected:
    virtual void showWindow() noexcept
    {
      runExcelThread([this]
      {
        LogWindow::showWindow();
      }, 0);
    }

    virtual void setWindowText() noexcept
    {
      runExcelThread([this]
      {
        LogWindow::setWindowText();
      }, 0);
    }
  };

  shared_ptr<ILogWindow> createLogWindow(
    HWND parentWindow,
    HINSTANCE parentInstance,
    const wchar_t* winTitle,
    HMENU menuBar,
    WNDPROC menuHandler,
    size_t historySize) noexcept
  {
    try
    {
      return make_shared<LogWindowThreaded>(
        parentWindow, parentInstance, winTitle, menuBar, menuHandler, historySize);
    }
    catch (...)
    {
      return shared_ptr<ILogWindow>();
    }
  }

  void loadFailureLogWindow(HINSTANCE parent, const std::wstring_view& msg) noexcept
  {
    // This function is always called from the main thread so can use
    // statics and the single-threaded version of LogWindow. 
    std::wstring msgStr;
    try
    {
      auto t = std::time(nullptr);
      tm tm;
      localtime_s(&tm, &t);
      msgStr = formatStr(L"%d-%d-%d: ", tm.tm_hour, tm.tm_min, tm.tm_sec).append(msg);

      static auto logWindow = make_shared<LogWindow>(
        (HWND)0, parent, L"xlOil Load Failure", (HMENU)0, (WNDPROC)0, 100);

      logWindow->appendMessage(std::move(msgStr));
      logWindow->openWindow();
    }
    catch (...)
    {
      OutputDebugString(msgStr.c_str());
    }
  }
}