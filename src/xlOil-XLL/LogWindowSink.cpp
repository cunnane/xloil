#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include "LogWindowSink.h"
#include <xloil/StringUtils.h>
#include <xlOilHelpers/Environment.h>
#include <xlOilHelpers/Exception.h>
#include <xlOil/LogWindow.h>
#include <spdlog/sinks/base_sink.h>
#include <spdlog/pattern_formatter.h>
#include <spdlog/spdlog.h>
#include <mutex>
#include <regex>


using std::wstring;
using std::string;
using std::shared_ptr;

namespace {
  constexpr short ID_CAPTURELEVEL = 400u;
  constexpr short ID_POPUPLEVEL = 500u;
  constexpr const wchar_t* LEVEL_NAMES[] = {
    L"Trace", L"Debug", L"Info", L"Warning", L"Error", L"Critical", L"Off"                                                                    \
  };
}

namespace xloil
{
  namespace MainLogWindow
  {
    HMENU theMenuBar;
    auto theCaptureLogLevel = spdlog::level::err;
    auto thePopupLevel      = spdlog::level::err;

    void setLogLevelRadioGroup(int base, int item)
    {
      CheckMenuRadioItem(
        theMenuBar,
        base,                     // first item in range 
        base + SPDLOG_LEVEL_OFF,  // last item in range 
        base + item,              // item to check 
        MF_BYCOMMAND              // IDs, not positions 
      );
    }

    void setCaptureLevel(spdlog::level::level_enum level)
    {
      setLogLevelRadioGroup(ID_CAPTURELEVEL, level);
      theCaptureLogLevel = level;
    }

    void setPopupLevel(spdlog::level::level_enum level)
    {
      setLogLevelRadioGroup(ID_POPUPLEVEL, level);
      thePopupLevel = level;
      if (theCaptureLogLevel > thePopupLevel)
        setCaptureLevel(level);
    }

    LRESULT CALLBACK MenuWindowProc(
      HWND hwnd,
      UINT message,
      WPARAM wParam,
      LPARAM lParam)
    {
      switch (message)
      {
      case WM_COMMAND:
      {
        const auto id = LOWORD(wParam);
        if (id >= ID_CAPTURELEVEL && id <= ID_CAPTURELEVEL + SPDLOG_LEVEL_OFF)
        {
          setCaptureLevel((spdlog::level::level_enum)(id - ID_CAPTURELEVEL));
        }
        else if (id >= ID_POPUPLEVEL && id <= ID_POPUPLEVEL + SPDLOG_LEVEL_OFF)
        {
          setPopupLevel((spdlog::level::level_enum)(id - ID_POPUPLEVEL));
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
      auto hCaptureMenu = CreateMenu();
      auto hPopupMenu = CreateMenu();
      for (short level = SPDLOG_LEVEL_OFF; level >= SPDLOG_LEVEL_TRACE; --level)
      {
        AppendMenu(hCaptureMenu, MF_STRING, ID_CAPTURELEVEL + level, LEVEL_NAMES[level]);
        AppendMenu(hPopupMenu,   MF_STRING, ID_POPUPLEVEL + level,   LEVEL_NAMES[level]);
      }

      AppendMenu(hMenubar, MF_POPUP, (UINT_PTR)hCaptureMenu, L"Capture Level");
      AppendMenu(hMenubar, MF_POPUP, (UINT_PTR)hPopupMenu,   L"Popup Level");

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
      if (msg.level < MainLogWindow::theCaptureLogLevel)
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

  void setLogWindowPopupLevel(const char* popupLevel)
  {
    const auto spdLevel = spdlog::level::from_str(popupLevel);
    MainLogWindow::setPopupLevel(spdLevel);
  }
}