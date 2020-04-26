#include "pch.h"
#include "CppUnitTest.h"
#include <xlOilHelpers/Environment.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;

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
      // TODO: how to make this into a test - i.e. what registry keys have
      // reliable values?
    }
  };
}