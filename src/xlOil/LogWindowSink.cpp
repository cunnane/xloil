#include "LogWindowSink.h"
#include <xloil/WindowsSlim.h>
#include <xloil/StringUtils.h>
#include <xlOilHelpers/Environment.h>
#include <xlOilHelpers/Exception.h>
#include <xlOil/LogWindow.h>
#include <spdlog/sinks/base_sink.h>
#include <spdlog/details/pattern_formatter.h>
#include <mutex>
#include <regex>

using std::wstring;
using std::string;
using std::shared_ptr;

namespace {
  constexpr unsigned ID_CAPTURELEVEL_ERROR = 40001;
  constexpr unsigned ID_CAPTURELEVEL_WARNING = 40002;
  constexpr unsigned ID_CAPTURELEVEL_INFO = 40003;
  constexpr unsigned ID_CAPTURELEVEL_DEBUG = 40004;
  constexpr unsigned ID_CAPTURELEVEL_TRACE = 40005;
}
namespace xloil
{
  namespace MainLogWindow
  {
    HMENU theMenuBar;
    spdlog::level::level_enum theSelectedLogLevel = spdlog::level::warn;
    spdlog::level::level_enum thePopupLevel = spdlog::level::err;

    LRESULT CALLBACK MenuWindowProc(
      HWND hwnd,
      UINT message,
      WPARAM wParam,
      LPARAM lParam)
    {
      constexpr int xOffset = 5, yOffset = 5;

      switch (message)
      {
      case WM_COMMAND:
      {
        const auto id = LOWORD(wParam);
        switch (id)
        {
        case ID_CAPTURELEVEL_ERROR:
        case ID_CAPTURELEVEL_WARNING:
        case ID_CAPTURELEVEL_INFO:
        case ID_CAPTURELEVEL_DEBUG:
        case ID_CAPTURELEVEL_TRACE:

          CheckMenuRadioItem(
            theMenuBar,
            ID_CAPTURELEVEL_ERROR,   // first item in range 
            ID_CAPTURELEVEL_TRACE,   // last item in range 
            id,                      // item to check 
            MF_BYCOMMAND             // IDs, not positions 
          );

          // Some sneaky but fragile enum arithmetic
          theSelectedLogLevel = (spdlog::level::level_enum)
            (spdlog::level::err - (id - ID_CAPTURELEVEL_ERROR));

          return 0;
        }
      }
      }

      return DefWindowProc(hwnd, message, wParam, lParam);
    }
  }

  class LogWindowSink : public spdlog::sinks::base_sink<std::mutex>
  {
  public:
    using level_enum = spdlog::level::level_enum;

    LogWindowSink(
      HWND parentWindow, 
      HINSTANCE parentInstance)
    {
      auto hMenubar = CreateMenu();
      auto hMenu = CreateMenu();

      AppendMenu(hMenu, MF_STRING,              ID_CAPTURELEVEL_ERROR,   L"Error");
      AppendMenu(hMenu, MF_STRING | MF_CHECKED, ID_CAPTURELEVEL_WARNING, L"Warning");
      AppendMenu(hMenu, MF_STRING,              ID_CAPTURELEVEL_INFO,    L"Info");
      AppendMenu(hMenu, MF_STRING,              ID_CAPTURELEVEL_DEBUG,   L"Debug");
      AppendMenu(hMenu, MF_STRING,              ID_CAPTURELEVEL_TRACE,   L"Trace");
      AppendMenu(hMenubar, MF_POPUP, (UINT_PTR)hMenu, L"Capture Level");

      MainLogWindow::theMenuBar = hMenubar;

      set_pattern_(""); // Just calls set_formatter_
      _window = createLogWindow(
        parentWindow, 
        parentInstance, 
        L"xlOil Log", 
        MainLogWindow::theMenuBar, 
        MainLogWindow::MenuWindowProc, 
        100);
    }

    void sink_it_(const spdlog::details::log_msg& msg) override
    {
      if (msg.level < MainLogWindow::theSelectedLogLevel)
        return;

      spdlog::memory_buf_t formatted;
      formatter_->format(msg, formatted);
      _window->appendMessage(
        utf8ToUtf16(
          std::string_view(formatted.data(), formatted.size())));

      if (msg.level >= MainLogWindow::thePopupLevel)
        _window->openWindow();
    }

    void flush_() override {}
    
    void set_formatter_(std::unique_ptr<spdlog::formatter> sink_formatter) override
    {
      base_sink::set_formatter_(std::make_unique<spdlog::pattern_formatter>(
        "[%H:%M:%S] [%l] %v"));
    }

    void openWindow()
    {
      _window->openWindow();
    }

  private:
    shared_ptr<ILogWindow> _window;
  };

  namespace
  {
    shared_ptr<LogWindowSink> theLogWindow;
  }

  std::shared_ptr<spdlog::sinks::sink>
    makeLogWindowSink(
      HWND parentWindow, 
      HINSTANCE parentInstance)
  {
    static auto ptr = new LogWindowSink(parentWindow, parentInstance);
    theLogWindow.reset(ptr);
    return theLogWindow;
  }

  void openLogWindow()
  {
    theLogWindow->openWindow();
  }

  void setLogWindowPopupLevel(spdlog::level::level_enum popupLevel)
  {
    MainLogWindow::thePopupLevel = popupLevel;
  }
}