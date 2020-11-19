#pragma once
#include <xloil/WindowsSlim.h>
#include <xloil/State.h>
#include <string>
#include <memory>

namespace xloil
{
  class ILogWindow
  {
  public:
    virtual void openWindow() noexcept = 0;
    virtual void appendMessage(const std::string& msg) noexcept = 0;
  };

  std::shared_ptr<ILogWindow> createLogWindow(
    HWND parentWindow,
    HINSTANCE parentInstance,
    const wchar_t* winTitle,
    HMENU menuBar,
    WNDPROC menuHandler,
    size_t historySize) noexcept;

  void writeLogWindow(const wchar_t* msg) noexcept;

  void writeLogWindow(const char* msg) noexcept;
}
