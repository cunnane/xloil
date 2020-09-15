#include "LogWindow.h"

#include <spdlog/sinks/base_sink.h>
#include <spdlog/details/pattern_formatter.h>
#include <xloil/WindowsSlim.h>
#include <xloil/StringUtils.h>
#include <mutex>
#include "resource.h"
#include <regex>

using std::wstring;
using std::string;

namespace xloil
{
  namespace LogWindow
  {
    HMENU theMenuBar;
    HWND theTextControl;
    HMENU theTextControlId = (HMENU)101;
    spdlog::level::level_enum theSelectedLogLevel = spdlog::level::warn;
    bool theWindowIsOpen = false;

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
        {
          /*auto err = GetLastErrorStdStr();
          MessageBox(NULL, TEXT("CreateWindow Failed!"), err.c_str(), MB_ICONERROR);*/
          return EXIT_FAILURE;
        }
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

        switch (LOWORD(wParam))
        {
        case ID_CAPTURELEVEL_ERROR:
          theSelectedLogLevel = spdlog::level::critical;
        case ID_CAPTURELEVEL_WARNING:
          theSelectedLogLevel = spdlog::level::warn;
        case ID_CAPTURELEVEL_INFO:
          theSelectedLogLevel = spdlog::level::info;
        case ID_CAPTURELEVEL_DEBUG:
          theSelectedLogLevel = spdlog::level::debug;
        case ID_CAPTURELEVEL_TRACE:
          theSelectedLogLevel = spdlog::level::trace;

          CheckMenuRadioItem(
            theMenuBar,
            ID_CAPTURELEVEL_ERROR, // first item in range 
            ID_CAPTURELEVEL_TRACE,    // last item in range 
            LOWORD(wParam),           // item to check 
            MF_BYCOMMAND              // IDs, not positions 
          );

          return 0;
        default:
          return DefWindowProc(hwnd, message, wParam, lParam);
        }

      default:
        return DefWindowProc(hwnd, message, wParam, lParam);
      }
    }

    void setTextBoxContents(const wchar_t* text)
    {
      // Add text to the window. 
      SendMessage(theTextControl, WM_SETTEXT, 0, (LPARAM)text);
    }

    HWND createWindow(HWND parentWnd, HINSTANCE hInstance, const wchar_t* winTitle)
    {
      // Define the main window class
      WNDCLASSEX win;
      win.cbSize = sizeof(WNDCLASSEX);
      win.hInstance = hInstance;
      win.lpszClassName = L"xlOil_Log";
      win.lpfnWndProc = WindowProc;
      win.style = CS_HREDRAW | CS_VREDRAW;

      // Use default icons and mouse pointer
      win.hIcon = LoadIcon(NULL, IDI_APPLICATION);
      win.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
      win.hCursor = LoadCursor(NULL, IDC_ARROW);

      win.lpszMenuName = NULL;
      win.cbClsExtra = 0;  // No extra bytes after the window class
      win.cbWndExtra = 0;  // No extra bytes after the the window instance
      win.hbrBackground = GetSysColorBrush(COLOR_3DFACE); // Use default colour

      // TODO: how to handle error?
      if (!RegisterClassEx(&win))
        return 0;

      theMenuBar = LoadMenu(hInstance, MAKEINTRESOURCE(IDR_LOG_WINDOW_MENU));

      /* The class is registered, let's create the program*/
      auto hwnd = CreateWindowEx(
        0, 
        win.lpszClassName,
        winTitle,
        WS_OVERLAPPEDWINDOW, // Title bar, minimimise, close and resize controls
        CW_USEDEFAULT, CW_USEDEFAULT, // (x, y)-position
        CW_USEDEFAULT, CW_USEDEFAULT, // width, height
        parentWnd,
        theMenuBar,
        hInstance,
        NULL // No Window Creation data
      );

      return hwnd;
    }

    void showWindow(HWND hwnd, size_t linesToScroll = 0)
    {
      ShowWindow(hwnd, SW_SHOWNORMAL);
      SendMessage(theTextControl, EM_LINESCROLL, 0, linesToScroll);
      theWindowIsOpen = true;
    }

  }

  class LogWindowSink : public spdlog::sinks::base_sink<std::mutex>
  {
  public:
    using level_enum = spdlog::level::level_enum;

    LogWindowSink(HWND parentWindow, HINSTANCE parentInstance)
    {
      _logWindow = LogWindow::createWindow(parentWindow, parentInstance, L"xlOil Log");
      set_pattern_(""); // Just calls set_formatter_
    }

    void sink_it_(const spdlog::details::log_msg& msg) override
    {
      if (msg.level < LogWindow::theSelectedLogLevel)
        return;

      spdlog::memory_buf_t formatted;
      formatter_->format(msg, formatted);
      _messages.emplace_back(fmt::to_string(formatted));

      if (_messages.size() > _maxSize)
        _messages.pop_front();

      if (isOpen())
      {
        appendToWindow(_messages.back());
        setWindowText();
      }
      else if (msg.level >= _popupLevel)
      {
        openWindow();
      }
    }

    void flush_() override {}
    
    void set_formatter_(std::unique_ptr<spdlog::formatter> sink_formatter) override
    {
      base_sink::set_formatter_(std::make_unique<spdlog::pattern_formatter>(
        "[%H:%M:%S] [%l] %v"));
    }
    
  private:
    bool isOpen() const
    {
      return LogWindow::theWindowIsOpen;
    }

    void openWindow()
    {
      _windowText.clear();
      for (auto& msg : _messages)
        appendToWindow(msg);
      
      setWindowText();

      LogWindow::showWindow(_logWindow, _messages.size());
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

    void setWindowText()
    {
      LogWindow::setTextBoxContents(_windowText.c_str());
    }

    std::list<string> _messages;
    wstring _windowText;
    size_t _maxSize = 100;
    level_enum _popupLevel = spdlog::level::err;
    HWND _logWindow;
  };

  std::shared_ptr<spdlog::sinks::sink>
    makeLogWindowSink(HWND parentWindow, HINSTANCE parentInstance)
  {
    return std::make_shared<LogWindowSink>(parentWindow, parentInstance);
  }
}