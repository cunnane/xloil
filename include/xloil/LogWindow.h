#pragma once
#include <xloil/State.h>
#include <string>
#include <memory>
#include <functional>

#ifndef _WINDOWS_
#include <basetsd.h>
#define DECLARE_HANDLE(name) struct name##__{int unused;}; typedef struct name##__ *name
DECLARE_HANDLE(HWND);
DECLARE_HANDLE(HINSTANCE);
DECLARE_HANDLE(HMENU);

#ifndef CALLBACK
#define CALLBACK __stdcall
#endif

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
    virtual void appendMessage(std::wstring&& msg) noexcept = 0;
  };

  std::shared_ptr<ILogWindow> createLogWindow(
    HWND parentWindow, // can be zero
    HINSTANCE parentInstance,
    const wchar_t* winTitle,
    HMENU menuBar,
    WNDPROC menuHandler,
    size_t historySize) noexcept;

  /// <summary>
  /// Called from the XLL initialisation code to report errors before the
  /// main logger has started.  Must be called on main thread - should be
  /// the case since is is called from AutoOpen.
  /// </summary>
  void loadFailureLogWindow(HINSTANCE parent, const std::wstring_view& msg, bool openWindow) noexcept;
}
