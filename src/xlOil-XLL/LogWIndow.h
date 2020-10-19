#pragma once
#include <xloil/WindowsSlim.h>
#include <xloil/State.h>
#include <string>
#include <memory>
namespace xloil
{
  namespace Helpers
  {
    class ILogWindow
    {
    public:
      virtual void openWindow() = 0;
      virtual void appendMessage(const std::string& msg) = 0;
    };

    std::shared_ptr<ILogWindow> createLogWindow(
      HWND parentWindow, 
      HINSTANCE parentInstance, 
      const wchar_t* winTitle,
      HMENU menuBar,
      WNDPROC menuHandler,
      size_t historySize);

    void writeLogWindow(const wchar_t* msg);

    void writeLogWindow(const char* msg);
  }
}
