#include "pch.h"
#include "CppUnitTest.h"
#include <xlOil/Date.h>

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

      dateTimeAddFormat(L"%Y%b%d");
      dateTimeAddFormat(L"%Y-%m-%d");

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
  };
}