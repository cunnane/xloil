#pragma once
#include <xloil/State.h>
#include <string>
#include <memory>

#ifndef _WINDOWS_
#include <basetsd.h>
#define DECLARE_HANDLE(name) struct name##__{int unused;}; typedef struct name##__ *name
DECLARE_HANDLE(HWND);
DECLARE_HANDLE(HINSTANCE);
DECLARE_HANDLE(HMENU);

typedef UINT_PTR WPARAM;
typedef LONG_PTR LPARAM;
typedef LONG_PTR LRESULT;
typedef LRESULT(CALLBACK* WNDPROC)(HWND, unsigned int, WPARAM, LPARAM);
#endif

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
