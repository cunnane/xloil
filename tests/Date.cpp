#include "CppUnitTest.h"
#include <xlOil/Date.h>
#include <xloil/ExcelObj.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;

namespace Tests
{
  TEST_CLASS(Date)
  {
  public:

    TEST_METHOD(Test_StrToDate)
    {
      std::tm result;
      bool ret;
      ret = stringToDateTime(L"2017-01-01", result, L"%Y-%m-%d");
      Assert::IsTrue(ret);

      theDateTimeFormats().push_back(L"%Y-%m-%d");
      theDateTimeFormats().push_back(L"%Y%b%d");

      ret = stringToDateTime(L"2010-02-03", result);
      Assert::IsTrue(ret);
      Assert::AreEqual(result.tm_year + 1900, 2010);
      Assert::AreEqual(result.tm_mon + 1, 2);
      Assert::AreEqual(result.tm_mday, 3);

      ret = stringToDateTime(L"2017Feb01", result);
      Assert::IsTrue(ret);
      Assert::AreEqual(result.tm_year + 1900, 2017);
      Assert::AreEqual(result.tm_mon + 1, 2);
      Assert::AreEqual(result.tm_mday, 1);
    }

    TEST_METHOD(Test_DateRoundTrip)
    {
      {
        int input[] = { 2020, 2, 1, 11, 30, 4, 516 }, *d = std::end(input);
        auto serial = excelSerialDateFromYMDHMS(*--d, *--d, *--d, *--d, *--d, *--d, *--d);
        int output[_countof(input)];
        d = std::end(output);
        excelSerialDatetoYMDHMS(serial, *--d, *--d, *--d, *--d, *--d, *--d, *--d);
        Assert::IsTrue(std::equal(input, std::end(input), output));
      }
    }

    void checkTMValues(const std::tm& tm, int year, int month, int day)
    {
      Assert::AreEqual(tm.tm_year + 1900, year);
      Assert::AreEqual(tm.tm_mon + 1, month);
      Assert::AreEqual(tm.tm_mday, day);
    }

    void testVisitorRoundTrip(const wchar_t* date, int year, int month, int day)
    {
      ExcelObj dateStr(date);
      ParseDateVisitor parseVisitor(L"%Y-%m-%d");
      auto ret = dateStr.visit(parseVisitor);
      Assert::IsTrue(ret);
      checkTMValues(parseVisitor.result, year, month, day);

      ExcelObj dateObj(
        excelSerialDateFromTM(parseVisitor.result, parseVisitor.uSecs));

      DateTimeVisitor dateVisitor;
      ret = dateObj.visit(dateVisitor);
      Assert::IsTrue(ret);
      checkTMValues(dateVisitor.result, year, month, day);
    }

    TEST_METHOD(Test_DateVisitors)
    {
      testVisitorRoundTrip(L"2017-01-01", 2017, 1, 1);
      testVisitorRoundTrip(L"1914-02-28", 1914, 2, 28);
    }
  };
}