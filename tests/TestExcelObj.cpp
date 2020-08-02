#include "CppUnitTest.h"
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelObj.h>

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

    TEST_METHOD(Comparison)
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
  };
}