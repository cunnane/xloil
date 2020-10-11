#include "CppUnitTest.h"
#include <xlOil/PString.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;

namespace Tests
{
  TEST_CLASS(PStringTests)
  {
  public:
    TEST_METHOD(Test_Create)
    {
      {
        PString str(L"Foo");
        Assert::IsTrue(str == L"Foo");
        str.resize(6);
        str = L"Foobar";
        Assert::IsTrue(str == L"Foobar");
      }
      {
        PString<> str(wstring(L"Foo"));
        Assert::IsTrue(str == L"Foo");
      }
      {
        PString<> str(3);
        str = wstring(L"Foo");
        Assert::IsTrue(str == L"Foo");
        // Check resize is automatic
        str = L"Foobar";
        Assert::IsTrue(str == L"Foobar");
      }
    }
    TEST_METHOD(Test_Strtok)
    {
      PString address(L"['My Book']'My Sheet'!A1");
      PStringView view(address);

      const auto delims = L"[]'!";
      const auto wb = view.strtok(delims);
      const auto ws = view.strtok(delims);
      const auto cell = view.strtok(delims);

      Assert::IsFalse(wb.empty());
      Assert::IsTrue(wb.view() == L"My Book");

      Assert::IsFalse(ws.empty());
      Assert::IsTrue(ws.view() == L"My Sheet");

      Assert::IsFalse(cell.empty());
      Assert::IsTrue(cell.view() == L"A1");
    }
    TEST_METHOD(Test_Append)
    {
      {
        PString str(L"Foo");
        auto sum = str + L"bar";
        Assert::IsTrue(sum == L"Foobar");
      }
      {
        PString str(L"Foo");
        auto sum = str + wstring(L"bar");
        Assert::IsTrue(sum == L"Foobar");
      }
      {
        PString str(L"Foo");
        auto sum = str + PString(L"bar");
        Assert::IsTrue(sum == L"Foobar");
      }
      {
        PString str(L"Foo");
        auto sum = wstring(L"Bar") + str;
        Assert::IsTrue(sum == L"BarFoo");
      }
      {
        PString str(L"Foo");
        auto sum = L"Bar" + str;
        Assert::IsTrue(sum == L"BarFoo");
      }
    }
  };
}