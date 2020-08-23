#include "CppUnitTest.h"
#include <xlOil/ExcelRef.h>
#include <xlOil/ExcelObj.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;

namespace Tests
{
  TEST_CLASS(TestRange)
  {
  public:

    TEST_METHOD(TestExcelRef)
    {
      msxll::IDSHEET sheet = nullptr;
      auto rng = newXllRange(ExcelRef(sheet, 1, 1, 10, 10));

      // We can only use local addresses as getting the sheet name requires an API call
      {
        auto subRange = unique_ptr<Range>(rng->range(1, 1, 2, 2));
        auto address = subRange->address(true);
        Assert::AreEqual(L"C3:D4", address.c_str());
      }
      {
        auto subRange = unique_ptr<Range>(rng->range(-1, -1, -1, -1));
        auto address = subRange->address(true);
        Assert::AreEqual(L"A1", address.c_str());
      }
      {
        auto subRange = unique_ptr<Range>(rng->range(2, 2));
        auto address = subRange->address(true);
        Assert::AreEqual(L"D4:K11", address.c_str());
      }
      {
        Assert::ExpectException<std::exception>([rng]() {
          rng->range(-2, -1);
        });
      }
    }
  };
}