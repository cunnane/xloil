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
    constexpr int MillisPerSecond = 1000;
    constexpr int MillisPerMinute = MillisPerSecond * 60;
    constexpr int MillisPerHour = MillisPerMinute * 60;
    constexpr int MillisPerDay = MillisPerHour * 24;
    constexpr auto millisecsPerDay = double(duration_cast<milliseconds>(hours(24)).count());
  }

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

  

  /*std::array<bool, 256> camelHerder()
  {
    std::array<bool, 256> result;
    for (auto c : "JFMASONDjfmasond")
      result[c] = true;
    return result;
  }

  static std::array<bool, 256> theFirstMonthLetters = camelHerder();

  void camel(wchar_t* str, size_t len)
  {
    const auto pEnd = str + len;
    while (!theFirstMonthLetters[(unsigned char)*str]) 
      if (++str == pEnd)
        return;
    *str = towupper(*str);
    while (++str < pEnd && iswalpha(*str))
      *str = towlower(*str);
  }*/

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