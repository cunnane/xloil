#include "CppUnitTest.h"
#include <xlOil/ExcelCall.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;

namespace Tests
{
  TEST_CLASS(TestExcelCall)
  {
  public:

    TEST_METHOD(TestExcelCallNameLookup)
    {
      Assert::AreEqual(excelFuncNumber("onWindow"), msxll::xlcOnWindow);
      Assert::AreEqual(excelFuncNumber("foobar"), -1);
      Assert::AreEqual(excelFuncNumber(excelFuncName(msxll::xlcOnWindow)), msxll::xlcOnWindow);
      Assert::AreEqual(excelFuncName(999), nullptr);
      Assert::AreEqual(excelFuncNumber("n"), msxll::xlfN);
      Assert::AreEqual(excelFuncNumber(excelFuncName(msxll::xlfT)), msxll::xlfT);
    }
  };
}