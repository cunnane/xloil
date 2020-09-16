#pragma once
#include <spdlog/spdlog.h>
namespace xloil
{
  std::shared_ptr<spdlog::sinks::sink>
    makeLogWindowSink(
      HWND parentWindow,
      HINSTANCE parentInstance,
      spdlog::level::level_enum popupLevel);
}