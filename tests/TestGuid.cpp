#include "CppUnitTest.h"
#include <xloil/WindowsSlim.h>
#include <xloilHelpers/GuidUtils.h>
#include <combaseapi.h> // StringFromGUID2

using namespace xloil;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;

const GUID theExcelDnaNamespaceGuid =
  { 0x306D016E, 0xCCE8, 0x4861, { 0x9D, 0xA1, 0x51, 0xA2, 0x7C, 0xBE, 0x34, 0x1A} };


namespace Tests
{
  TEST_CLASS(TestStableGuid)
  {
  public:

    TEST_METHOD(TestExcelDnaGuid)
    {
      // The expected results are calculated by running this code:
      // https://github.com/Excel-DNA/ExcelDna/blob/661aba734f08b2537866632f2295a77062640672/Source/ExcelDna.Integration/GuidUtility.cs
      // on dotnetfiddle.net
      wchar_t str[64];
      GUID result;
      {
        stableGuidFromString(result, theExcelDnaNamespaceGuid, L"");
        StringFromGUID2(result, str, _countof(str));
        Assert::AreEqual(L"{FB475EF6-2275-54B0-AF78-FE136DAF4ECA}", str);
      }
      {
        stableGuidFromString(result, theExcelDnaNamespaceGuid, L"addin.xll");
        StringFromGUID2(result, str, _countof(str));
        Assert::AreEqual(L"{818CECFE-DB3F-51CF-92A3-22D1727F623D}", str);
      }
      {
        stableGuidFromString(result, theExcelDnaNamespaceGuid, L"c:\\path\\addin.xll");
        StringFromGUID2(result, str, _countof(str));
        Assert::AreEqual(L"{02DCD9CA-1AB0-5F79-97A0-81A3DBBD267F}", str);
      }
    }
  };
}
