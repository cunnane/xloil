#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>

#include "LogWindow.h"

#include <xloil/StringUtils.h>
#include <xlOilHelpers/Environment.h>
#include <xlOilHelpers/Exception.h>
#include <list>

using std::wstring;
using std::string;

namespace xloil 
{
  namespace Helpers
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
      const wchar_t* theWindowClass = L"xlOil_Log";

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
        static auto win = createWindowClass(hInstance, theWindowClass);

        // If we try to create a window from our class during xlAutoOpen, we will get
        // the cryptic "ntdll.dll (EXCEL.EXE) RangeChecks instrumentation code detected 
        // an out of range array access".  Whatever Excel gets up to during start-up
        // seems to screw around with the Win32 API.  We use the existence or 
        // otherwise of a parentHwnd to determine if we are in this perilous state

        auto hwnd = CreateWindowEx(
          0,
          parentWnd ? MAKEINTATOM(win) : L"EDIT",
          winTitle,
          parentWnd 
            ? WS_OVERLAPPEDWINDOW  // Title bar, minimimise, close and resize controls
            : WS_OVERLAPPEDWINDOW | WS_VSCROLL | ES_LEFT | ES_MULTILINE,
          CW_USEDEFAULT, CW_USEDEFAULT, // (x, y)-position
          CW_USEDEFAULT, CW_USEDEFAULT, // width, height
          HWND_DESKTOP,
          menuBar,
          hInstance,
          this
        );

        if (!hwnd)
          throw Exception(L"Failed to create LogWindow: %s",
            writeWindowsError().c_str());

        if (!parentWnd)
          theTextControl = hwnd;

        theMenuHandler = menuHandler;
        theMainWindow = hwnd;
      }

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
            (HINSTANCE)GetWindowLong(hwnd, GWLP_HINSTANCE),
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

      static ATOM createWindowClass(HINSTANCE hInstance, const wchar_t* windowClass)
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

        auto res = RegisterClassEx(&win);
        if (!res)
          throw Exception(L"Failed to create LogWindow: %s", 
            writeWindowsError().c_str());

        return res;
      }

      void showWindow()
      {
        ShowWindow(theMainWindow, SW_SHOWNORMAL);
        theWindowIsOpen = true;
      }

      void setWindowText()
      {
        // Just scroll to the end, word-wrap seems to confuse the
        // line count, so we just specify a big number
        setTextBoxContents(_windowText.c_str(), 66666666);
      }

      bool isOpen() const
      {
        return theWindowIsOpen;
      }

      void openWindow() override
      {
        if (isOpen())
          return;

        _windowText.clear();
        for (auto& msg : _messages)
          appendToWindow(msg);

        setWindowText();

        showWindow();
      }

      void appendToWindow(const string& msg)
      {
        auto wmsg = utf8ToUtf16(msg);

        // Fix any unix line endings (e.g. from python)
        size_t pos = -1;
        while ((pos = wmsg.find(L'\n', pos + 1)) != string::npos)
        {
          if (wmsg[pos - 1] != L'\r')
            wmsg.replace(pos, 2, L"\r\n");
        }

        _windowText.append(wmsg).append(L"\r\n");
      }

      void appendMessage(const string& msg) override
      {
        _messages.push_back(msg);

        if (_messages.size() > _maxSize)
          _messages.pop_front();


        if (isOpen())
        {
          appendToWindow(_messages.back());
          setWindowText();
        }
      }
    };

    std::shared_ptr<ILogWindow> createLogWindow(
      HWND parentWindow,
      HINSTANCE parentInstance,
      const wchar_t* winTitle,
      HMENU menuBar,
      WNDPROC menuHandler,
      size_t historySize)
    {
      return std::make_shared<LogWindow>(
        parentWindow, parentInstance, winTitle, menuBar, menuHandler, historySize);
    }
  }
}