#pragma once
#include "ExportMacro.h"

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
    excelSerialDatetoDMYHMS(double serial, int &nDay, int &nMonth, int &nYear, int& nHours, int& nMins, int& nSecs, int& uSecs);

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
    excelSerialDateFromDMYHMS(int nDay, int nMonth, int nYear, int nHours, int nMins, int nSecs, int uSecs);
}
