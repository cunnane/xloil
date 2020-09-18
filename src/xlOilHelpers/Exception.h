#pragma once

#include <xloil/StringUtils.h>
#include <stdexcept>

namespace xloil
{
  namespace Helpers
  {
    /// <summary>
    /// Lightweight exception class for use in xloilHelpers
    /// </summary>
    class Exception : public std::runtime_error
    {
    public:
      template<class... Args>
      Exception(
        const char* str, Args &&... args)
        : Exception(formatStr(str, std::forward<Args>(args)...))
      {}

      template<class... Args>
      Exception(
        const wchar_t* str, Args &&... args)
        : Exception(formatStr(str, std::forward<Args>(args)...))
      {}

      inline Exception(const std::wstring& msg)
        : runtime_error(utf16ToUtf8(msg.c_str()))
      {}

      inline Exception(
        const std::string& msg)
        : runtime_error(msg.c_str())
      {}
    };

    std::wstring writeWindowsError();
  }
}