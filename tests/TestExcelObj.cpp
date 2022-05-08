#include "CppUnitTest.h"
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>
#include <xlOil/Date.h>

#include <vector>


using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::vector;
using std::hash;

namespace Tests
{
  TEST_CLASS(TestExcelObj)
  {
  public:

    TEST_METHOD(TestComparison)
    {
      Assert::IsFalse(ExcelObj(L"Hello") < ExcelObj(L"hello"));
      Assert::IsTrue(ExcelObj(1.5) < ExcelObj(2));
      Assert::IsTrue(hash<ExcelObj>()(ExcelObj(1.5))
        == hash<ExcelObj>()(ExcelObj(1.5)));

      vector<double> smallerValues = { 1, 2, 3, 4 };
      auto smaller1 = ExcelObj(smallerValues.begin(), smallerValues.end());
      auto smaller2 = ExcelObj(smallerValues.begin(), smallerValues.end());

      Assert::IsTrue(smaller1 == smaller2);

      vector<double> largerValues = { 1, 2, 3, 5 };
      Assert::AreEqual(-1, ExcelObj::compare(
        smaller1,
        ExcelObj(largerValues.begin(), largerValues.end()),
        true, true));
    }

    TEST_METHOD(TestStrings)
    {
      {
        Assert::IsTrue(ExcelObj(L"Foo") == L"Foo");
        Assert::IsTrue(ExcelObj(L"Foo") == wstring(L"Foo"));
        Assert::IsTrue(ExcelObj("Foo") == L"Foo");
        Assert::IsTrue(ExcelObj("Foo") == wstring(L"Foo"));
        Assert::IsTrue(ExcelObj("") == L"");
        Assert::IsFalse(ExcelObj(3) == L"Foo");
      }
    }

    TEST_METHOD(TestArray)
    {
      ExcelObj obj = { 1, 2, 3, 4, 5 };
      ExcelArray arr(obj);
      Assert::IsTrue(arr(0) == 1);
      Assert::IsTrue(arr(2) == 3);
      Assert::IsTrue(arr(4) == 5);
      vector<double> values;
      std::transform(arr.begin(), arr.end(), std::back_inserter(values), ApplyVisitor{ conv::ToType<double>() });
      Assert::IsTrue(values == vector<double>{ 1, 2, 3, 4, 5 });
    }

    TEST_METHOD(TestArraySlice)
    {
      ExcelObj obj = { { 1, 2, 3 }, { 4, 5, 6} };
      ExcelArray arr(obj);

      {
        auto negativeSlice = arr.slice(0, -2);
        Assert::AreEqual(arr(0, 1).get<int>(), negativeSlice(0, 0).get<int>());
      }

      {
        auto slice2d = arr.slice(0, 1, 2, 3);
        Assert::AreEqual(arr(0, 1).get<int>(), slice2d(0, 0).get<int>());
        Assert::AreEqual(2u, slice2d.nRows());
        Assert::AreEqual(2u, slice2d.nCols());
      }

      // Out-of-bounds slice
      Assert::ExpectException<std::out_of_range>([&]() { arr.slice(3, 0, 5, 0); });
    }
    TEST_METHOD(TestCreateFromDate)
    {
      {
        ExcelObj date(stringToDateTime(L"2017-01-01", L"%Y-%m-%d"));
        auto tm = date.get<std::tm>();
        Assert::AreEqual(2017, tm.tm_year + 1900);
        Assert::AreEqual(1, tm.tm_mon + 1);
        Assert::AreEqual(1, tm.tm_mday);
      }
    }
  };
}