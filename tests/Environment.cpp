#include "CppUnitTest.h"
#include <xlOilHelpers/Environment.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::string;
namespace Tests
{
  TEST_CLASS(Environment)
  {
  public:
    TEST_METHOD(Test_expandWindowsRegistryStrings)
    {
      wstring test(L"xls file: <HKCR\\.xls\\>, "
        "username: <HKCU\\Software\\Microsoft\\Office\\Common\\UserInfo\\UserName>. Done.");

      auto result = expandWindowsRegistryStrings(test);
      // TODO: how to make this into a test - i.e. what registry keys have known values?
    }

    TEST_METHOD(Test_expandEnvironmentStrings)
    {
      wstring test(L"xls file: <HKCR\\.xls\\>, "
        "username: <HKCU\\Software\\Microsoft\\Office\\Common\\UserInfo\\UserName>. Done.");
      {
        auto result = getEnvVar("TEMP");
        string expected;
        size_t len = result.length() + 1;
        expected.resize(len);
        Assert::AreEqual(0, getenv_s(&len, expected.data(), len, "TEMP"));
        Assert::AreEqual(len, result.length() + 1);
        expected.pop_back(); // Remove null terminator
        Assert::AreEqual(expected, result);
      }
      {
        auto expected = getEnvVar(L"TEMP");
        auto result = expandEnvironmentStrings(L"%TEMP%");
        Assert::AreEqual(expected, result);
      }
    }
  };
}