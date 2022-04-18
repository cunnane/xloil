#include <xlOil/StringUtils.h>
#include <xlOil/FmtStr.h>
#include "ExportMacro.h"

#pragma once

/// <summary>
/// Throws an xloil::Exception. Accepts python format strings like the logging functions
/// e.g. XLO_THROW("Bad: {0} {1}", errCode, e.what()) 
/// </summary>
#define XLO_THROW(...) do { throw xloil::Exception<>(__FILE__, __LINE__, __FUNCTION__, XLO_FMT_((__VA_ARGS__))); } while(false)
#define XLO_ASSERT(condition) assert((condition) && #condition)
#define XLO_THROW_TYPE(Type, ...) do { throw xloil::Exception<Type>(__FILE__, __LINE__, __FUNCTION__, XLO_FMT_((__VA_ARGS__))); } while(false)

namespace xloil
{
  /// <summary>
  /// Wrapper around GetLastError and FormatMessage to write out any error condition
  /// set by Windows API functions.
  /// </summary>
  XLOIL_EXPORT std::wstring writeWindowsError();

  XLOIL_EXPORT void logException(
    const char* path,
    const int line,
    const char* func,
    const char* msg) noexcept;

  namespace detail
  {
    inline std::string toUtf8(std::string&& str) { return std::move(str); }
    inline std::string toUtf8(std::wstring&& wstr) { return utf16ToUtf8(wstr); }

  } // detail

#pragma warning(disable: 4275) // Complaints about dll-interface. MS suggests disabling
  template<class TBase = std::runtime_error>
  class Exception : public TBase
  {
  public:
    template<class TMsg, class... Args>
    Exception(
      const char* path, const int line, const char* func, TMsg&& msg)
      : TBase(detail::toUtf8(std::forward<TMsg>(msg)))
    {
      logException(path, line, func, TBase::what());
    }
  };
}