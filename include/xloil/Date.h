#pragma once
#include "ExportMacro.h"
#include <xloil/ExcelObj.h>
#include <string_view>
#include <time.h>
#include <vector>

namespace std { struct tm; }
namespace xloil
{
  constexpr int XL_MAX_SERIAL_DATE = 2958465; // 31 December 9999

  /// <summary>
  /// Converts as Excel date expressed as an integer to day, month, year
  /// Returns true if the conversion was successful, otherwise the int
  /// is out of range of valid Excel dates.
  /// </summary>
  XLOIL_EXPORT bool 
    excelSerialDateToYMD(int nSerialDate, int &nYear, int &nMonth, int &nDay) noexcept;

  /// <summary>
  /// Converts as Excel date expressed as floating point to day, month, year,
  /// hours, minutes, seconds and milliseconds.
  /// Returns true if the conversion was successful, otherwise the value
  /// is out of range of valid Excel dates.
  /// </summary>
  XLOIL_EXPORT bool
    excelSerialDatetoYMDHMS(
      double serial, int &nYear, int &nMonth, int &nDay,
       int& nHours, int& nMins, int& nSecs, int& uSecs) noexcept;

  /// <summary>
  /// Converts a date specifed as day, month, year to an Excel date serial number
  /// </summary>
  XLOIL_EXPORT int
    excelSerialDateFromYMD(int nYear, int nMonth, int nDay) noexcept;

  /// <summary>
  /// Converts a date specifed as day, month, year, hours, minutes, seconds and 
  /// milliseconds to an Excel date serial number
  /// </summary>
  XLOIL_EXPORT double
    excelSerialDateFromYMDHMS(
      int nYear, int nMonth, int nDay,
      int nHours, int nMins, int nSecs, int uSecs) noexcept;

  inline double
    excelSerialDateFromTM(std::tm& tm, int uSecs = 0) noexcept
  {
    return excelSerialDateFromYMDHMS(tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday,
      tm.tm_hour, tm.tm_min, tm.tm_sec, uSecs);
  }

  /// <summary>
  /// Parses a string into a std::tm struct. Note that in
  /// the tm struct, the fields do not have the "most natural"
  /// bases: years are since 1900 and months start at zero.
  /// 
  /// If `format` is omitted, tries to parse the date using
  /// all registered formats <see cref="dateTimeAddFormat"/>.
  /// 
  /// Because it uses `std::get_time`, month matching is 
  /// case sensitive based on the convention in the current
  /// locale. This is such a painful restriction that replacing
  /// the parsing with a better designed library is an open 
  /// issue (unfortunately there is nothing particularly 
  /// lightweight available).
  /// </summary>
  XLOIL_EXPORT bool stringToDateTime(
    const std::wstring_view& str,
    std::tm& result, 
    const wchar_t* format = nullptr);

  inline std::tm stringToDateTime(
    const std::wstring_view& str,
    const wchar_t* format)
  {
    std::tm result;
    stringToDateTime(str, result, format);
    return result;
  }

  /// <summary>
  /// Registers date time formats to try when parsing strings with
  /// <see cref="stringToDateTime"/>.  See `std::get_time` for format syntax.
  /// </summary>
  XLOIL_EXPORT std::vector<std::wstring>& theDateTimeFormats();

  struct DateVisitor : public ExcelValVisitor<bool>
  {
    std::tm result = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };

    template <class T> bool operator()(T)
    {
      return false;
    }

    bool operator()(int x)
    {
      if (!excelSerialDateToYMD(x,
        result.tm_year, result.tm_mon, result.tm_mday))
        return false;
      result.tm_year -= 1900;
      result.tm_mon -= 1;
      return true;
    }

    bool operator()(double x)
    {
      return operator()((int)x); // Truncate or require exact?
    }
  };

  struct DateTimeVisitor : public DateVisitor
  {
    int uSecs = 0;

    using DateVisitor::operator();

    bool operator()(double x)
    {
      if (!excelSerialDatetoYMDHMS(x,
        result.tm_year, result.tm_mon, result.tm_mday,
        result.tm_hour, result.tm_min, result.tm_sec, uSecs))
        return false;
      result.tm_year -= 1900;
      result.tm_mon -= 1;
      return true;
    }
  };

  struct ParseDateVisitor : public DateTimeVisitor
  {
    ParseDateVisitor(const wchar_t* fmt = nullptr) 
      : format(fmt) {}

    const wchar_t* format;

    using DateVisitor::operator();
    bool operator()(PStringRef str)
    {
      return stringToDateTime(str, result, format);
    }
  };


  template <>
  inline std::tm ExcelObj::get() const
  {
    DateTimeVisitor v;
    if (visit(v))
      return std::move(v.result);
    XLO_THROW("Could not convert to datetime");
  }
}
