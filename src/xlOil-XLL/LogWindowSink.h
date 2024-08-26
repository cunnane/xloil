#pragma once
#include <memory>
#include <xloil/WindowsSlim.h>
namespace spdlog { namespace sinks { class sink; } }

namespace xloil
{
  std::shared_ptr<spdlog::sinks::sink>
    makeLogWindowSink(
      HWND parentWindow,
      HINSTANCE parentInstance);

  void openLogWindow();
  void setLogWindowPopupLevel(const char* popupLevel);
}