#include <xlOil/Date.h>
#include <cmath>
#include <chrono>
#include <streambuf>
#include <istream>
#include <iomanip>
#include <unordered_set>

using namespace std::chrono;
using std::unordered_set;
using std::string;
using std::wstring;

namespace xloil
{
  namespace
  {
    constexpr auto microsecsPerDay = double(duration_cast<microseconds>(hours(24)).count());
  }

  /// Verbatim from https://www.codeproject.com/Articles/2750/Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa
  bool excelSerialDateToYMD(int nSerialDate, int &nYear, int &nMonth, int &nDay) noexcept
  {
    if (nSerialDate > XL_MAX_SERIAL_DATE || nSerialDate < 0)
      return false;

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
    return true;
  }

  bool excelSerialDatetoYMDHMS(
    double serial, int &nYear, int &nMonth, int &nDay, int& nHours, int& nMins, int& nSecs, int& uSecs) noexcept
  {
    // Allow for time component in max date
    if (serial > (double)(XL_MAX_SERIAL_DATE + 1) || serial < 0) 
      return false;

    double intpart;
    if (std::modf(serial, &intpart) != 0.0)
    {
      auto us = microseconds(long long((serial - intpart) * microsecsPerDay));
      auto secs = duration_cast<seconds>(us);
      us -= duration_cast<microseconds>(secs);
      auto mins = duration_cast<minutes>(secs);
      secs -= duration_cast<seconds>(mins);
      auto hour = duration_cast<hours>(mins);
      mins -= duration_cast<minutes>(hour);

      nHours = hour.count();
      nMins = mins.count();
      nSecs = (int)secs.count();
      uSecs = (int)us.count();
    }
    else
      nHours = nMins = nSecs = uSecs = 0;

    return excelSerialDateToYMD(int(intpart), nYear, nMonth, nDay);
  }

  /// Verbatim from https://www.codeproject.com/Articles/2750/Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa
  int excelSerialDateFromYMD(int nYear, int nMonth, int nDay) noexcept
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

  double excelSerialDateFromYMDHMS(
    int nYear, int nMonth, int nDay, int nHours, int nMins, int nSecs, int uSecs) noexcept
  {
    const auto micros = duration_cast<microseconds>(
      hours(nHours) + minutes(nMins) + seconds(nSecs)).count() + uSecs;
    const auto serial = excelSerialDateFromYMD(nYear, nMonth, nDay) 
      + micros / microsecsPerDay;
    return serial;
  }

  // Thanks to:
  // https://stackoverflow.com/questions/13059091/creating-an-input-stream-from-constant-memory/13059195#13059195
  struct wmembuf : std::wstreambuf 
  {
    wmembuf(wchar_t const* base, size_t size) 
    {
      str(base, size);
    }
    void str(wchar_t const* base, size_t size)
    {
      wchar_t* p = const_cast<wchar_t*>(base);
      this->setg(p, p, p + size);
    }
  };
  struct wimemstream : virtual wmembuf, std::wistream
  {
    using std::wistream::imbue;
    wimemstream(wchar_t const* base, size_t size)
      : wmembuf(base, size)
      , std::wistream(static_cast<std::wstreambuf*>(this))
    {}
    void str(wchar_t const* base, size_t size)
    {
      wmembuf::str(base, size);
    }
  };

  unordered_set<wstring> theDateFormats;

  bool stringToDateTime(
    const std::wstring_view& str,
    std::tm& result, 
    const wchar_t* format)
  {
    wimemstream stream(str.data(), str.length());
    memset(&result, 0, sizeof(std::tm));

    if (format)
    {
      stream >> std::get_time(&result, format);
      return !stream.fail();
    }
    else
    {
      for (auto& form : theDateFormats)
      {
        stream >> std::get_time(&result, form.c_str());
        if (!stream.fail())
          return true;
        stream.clear();
        stream.str(str.data(), str.length());
      }
      return false;
    }
  }

  void dateTimeAddFormat(const wchar_t* f)
  {
    theDateFormats.insert(f);
  }
}