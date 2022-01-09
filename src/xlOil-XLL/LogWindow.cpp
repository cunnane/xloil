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
      HMENU theMenuBar;
      HWND theMainWindow;
      HWND theTextControl;
      HMENU theTextControlId = (HMENU)101;
      WNDPROC theMenuHandler;
      bool theWindowIsOpen = false;
      std::list<string> _messages;
      wstring _windowText;
      size_t _maxSize;
      shared_ptr<ATOM> _windowClass;
      static constexpr wchar_t* theWindowClass = L"xlOil_Log";

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

      void setTextBoxContents(const wchar_t* text, size_t numLines = 0)
      {
        // Add text to the window. 
        SendMessage(theTextControl, WM_SETTEXT, 0, (LPARAM)text);
        // Scroll to the last line
        SendMessage(theTextControl, EM_LINESCROLL, 0, numLines);
      }

      void showWindow() noexcept
      {
        ShowWindow(theMainWindow, SW_SHOWNORMAL);
        theWindowIsOpen = true;
      }

      void setWindowText() noexcept
      {
        // Just scroll to the end, word-wrap seems to confuse the
        // line count, so we just specify a big number
        setTextBoxContents(_windowText.c_str(), 66666666);
      }

      bool isOpen() const
      {
        return theWindowIsOpen;
      }

      void appendToWindow(const string& msg) 
      {
        auto wmsg = utf8ToUtf16(msg);

        // Fix any unix line endings (e.g. from python)
        auto pos = (size_t) -1;
        while ((pos = wmsg.find(L'\n', pos + 1)) != string::npos)
        {
          if (wmsg[pos - 1] != L'\r')
            wmsg.replace(pos, 1, L"\r\n");
        }

        _windowText.append(wmsg).append(L"\r\n");
      }

    public:
      void openWindow() noexcept override
      {
        if (isOpen())
          return;

        _windowText.clear();
        for (auto& msg : _messages)
          appendToWindow(msg);

        // Should only call window drawing API functions on the main thread
        runExcelThread([this]
        {
          setWindowText();
          showWindow();
        }, 0);
      }

      void appendMessage(const string& msg) noexcept override
      {
        _messages.push_back(msg);

        if (_messages.size() > _maxSize)
          _messages.pop_front();


        if (isOpen())
        {
          appendToWindow(_messages.back());
          // Should only call window drawing API functions on the main thread
          runExcelThread([this]
          {
            setWindowText();
          }, 0);
        }
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
        return make_shared<LogWindow>(
          parentWindow, parentInstance, winTitle, menuBar, menuHandler, historySize);
      }
      catch (...)
      {
        return shared_ptr<ILogWindow>();
      }
    }

    void writeLogWindow(const wchar_t* msg) noexcept
    {
      writeLogWindow(utf16ToUtf8(msg).c_str());
    }

    void writeLogWindow(const char* msg) noexcept
    {
      // Thread safe since C++11
      static auto logWindow = createLogWindow(
        0, (HINSTANCE)State::coreModuleHandle(), L"xlOil Load Failure", 0, 0, 100);

      if (!msg || !logWindow)
        return;

      auto t = std::time(nullptr);
      tm tm;
      localtime_s(&tm, &t);
      logWindow->appendMessage(
        formatStr("%d-%d-%d: %s", tm.tm_hour, tm.tm_min, tm.tm_sec, msg));
      logWindow->openWindow();
    }
}