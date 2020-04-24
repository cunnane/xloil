#include "Date.h"
#include <cmath>
#include <chrono>
using namespace std::chrono;

namespace xloil
{
  const int MillisPerSecond = 1000;
  const int MillisPerMinute = MillisPerSecond * 60;
  const int MillisPerHour = MillisPerMinute * 60;
  const int MillisPerDay = MillisPerHour * 24;

  /// Verbatim from https://www.codeproject.com/Articles/2750/Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa
  bool excelSerialDateToDMY(int nSerialDate, int &nDay, int &nMonth, int &nYear)
  {
    // TODO: range check???

    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (nSerialDate == 60)
    {
      nDay = 29;
      nMonth = 2;
      nYear = 1900;

      return true;
    }
    else if (nSerialDate < 60)
    {
      // Because of the 29-02-1900 bug, any serial date 
      // under 60 is one off... Compensate.
      nSerialDate++;
    }

    // Modified Julian to DMY calculation with an addition of 2415019
    int l = nSerialDate + 68569 + 2415019;
    int n = int((4 * l) / 146097);
    l = l - int((146097 * n + 3) / 4);
    int i = int((4000 * (l + 1)) / 1461001);
    l = l - int((1461 * i) / 4) + 31;
    int j = int((80 * l) / 2447);
    nDay = l - int((2447 * j) / 80);
    l = int(j / 11);
    nMonth = j + 2 - (12 * l);
    nYear = 100 * (n - 49) + i + l;
    return false;
  }

  constexpr auto millisecsPerDay = double(duration_cast<milliseconds>(hours(24)).count());

  bool excelSerialDatetoDMYHMS(
    double serial, int &nDay, int &nMonth, int &nYear, int& nHours, int& nMins, int& nSecs, int& uSecs)
  {
    double intpart;
    if (std::modf(serial, &intpart) != 0.0)
    {
      auto ms = milliseconds(long((serial - intpart) * millisecsPerDay));
      auto secs = duration_cast<seconds>(ms);
      ms -= duration_cast<milliseconds>(secs);
      auto mins = duration_cast<minutes>(secs);
      secs -= duration_cast<seconds>(mins);
      auto hour = duration_cast<hours>(mins);
      mins -= duration_cast<minutes>(hour);

      nHours = hour.count();
      nMins = mins.count();
      nSecs = (int)secs.count();
      uSecs = (int)ms.count();
    }
    else
      nHours = nMins = nSecs = uSecs = 0;
    excelSerialDateToDMY(int(intpart), nDay, nMonth, nYear);
    return true;
  }

  /// Verbatim from https://www.codeproject.com/Articles/2750/Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa
  int excelSerialDateFromDMY(int nDay, int nMonth, int nYear)
  {
    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (nDay == 29 && nMonth == 02 && nYear == 1900)
      return 60;

    // DMY to Modified Julian calculated with an extra subtraction of 2415019.
    long nSerialDate =
      int((1461 * (nYear + 4800 + int((nMonth - 14) / 12))) / 4) +
      int((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
      int((3 * (int((nYear + 4900 + int((nMonth - 14) / 12)) / 100))) / 4) +
      nDay - 2415019 - 32075;

    if (nSerialDate < 60)
    {
      // Because of the 29-02-1900 bug, any serial date 
      // under 60 is one off... Compensate.
      nSerialDate--;
    }

    return (int)nSerialDate;
  }

  double excelSerialDateFromDMYHMS(int nDay, int nMonth, int nYear, int nHours, int nMins, int nSecs, int uSecs)
  {
    double serial = excelSerialDateFromDMY(nDay, nMonth, nYear);
    auto ms = duration_cast<milliseconds>(hours(nHours) + minutes(nMins) + seconds(nSecs)).count() + uSecs;
    serial += ms / millisecsPerDay;
    return serial;
  }
}