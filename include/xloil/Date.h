#pragma once
#include "ExportMacro.h"
#include <string_view>

namespace std { struct tm; }
namespace xloil
{
  /// <summary>
  /// Converts as Excel date expressed as an integer to day, month, year
  /// Returns true if the conversion was successful, otherwise the int
  /// is out of range of valid Excel dates.
  /// </summary>
  XLOIL_EXPORT bool 
    excelSerialDateToDMY(int nSerialDate, int &nDay, int &nMonth, int &nYear);

  /// <summary>
  /// Converts as Excel date expressed as floating point to day, month, year,
  /// hours, minutes, seconds and milliseconds.
  /// Returns true if the conversion was successful, otherwise the value
  /// is out of range of valid Excel dates.
  /// </summary>
  XLOIL_EXPORT bool
    excelSerialDatetoDMYHMS(
      double serial, int &nDay, int &nMonth, 
      int &nYear, int& nHours, int& nMins, int& nSecs, int& uSecs);

  /// <summary>
  /// Converts a date specifed as day, month, year to an Excel date serial number
  /// </summary>
  XLOIL_EXPORT int
    excelSerialDateFromDMY(int nDay, int nMonth, int nYear);

  /// <summary>
  /// Converts a date specifed as day, month, year, hours, minutes, seconds and milliseconds
  /// to an Excel date serial number
  /// </summary>
  XLOIL_EXPORT double
    excelSerialDateFromDMYHMS(
      int nDay, int nMonth, int nYear, 
      int nHours, int nMins, int nSecs, int uSecs);

  /// <summary>
  /// Parses a string into a std::tm struct. Note that in
  /// the m struct, the fields do not have the "most natural"
  /// bases: years are since 1900 and months start at zero.
  /// 
  /// If `format` is omitted, tries to parse the date using
  /// all registered formats <see cref="dateTimeAddFormat"/>.
  /// 
  /// Because it uses std::get_time, month matching is 
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

  XLOIL_EXPORT void dateTimeAddFormat(const wchar_t* f);
}
