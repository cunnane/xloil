#include "CppUnitTest.h"
#include <xlOil/ExcelRef.h>
#include <xlOil/ExcelObj.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;
using std::to_string;

bool operator==(const msxll::XLREF12& l, const msxll::XLREF12& r)
{
  return memcmp(&l, &r, sizeof(l)) == 0;
}

namespace Microsoft {
  namespace VisualStudio {
    namespace CppUnitTestFramework
    {
      template<> inline std::wstring ToString<msxll::XLREF12>(const msxll::XLREF12& r) 
      { 
        return fmt::format(L"(row=[{},{}], col=[{},{}])", r.rwFirst, r.rwLast, r.colFirst, r.colLast);
      }
    }
  }
}

namespace Tests
{
  TEST_CLASS(TestRange)
  {
  public:

    TEST_METHOD(TestExcelRef)
    {
      msxll::IDSHEET sheet = nullptr;
      auto rng = new XllRange(ExcelRef(sheet, 1, 1, 10, 10));

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

    void xlRefRoundTrip(const msxll::XLREF12& ref, bool lowerCase = false)
    {
      wchar_t address[XL_FULL_ADDRESS_A1_MAX_LEN];
      xlrefToLocalA1(ref, address, _countof(address));
      if (lowerCase)
        std::transform(address, address + _countof(address), address, towlower);
      msxll::XLREF12 outRef;
      localAddressToXlRef(outRef, address);
      Assert::AreEqual(ref, outRef);
    }
    TEST_METHOD(TestAddressParseA1)
    {
      for (auto i = 1; i < XL_MAX_ROWS; i += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, i, 1, 1 });
      }
      for (auto j = 1; j < XL_MAX_COLS; j += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, 1, j, j });
      }
      for (auto j = 1; j < XL_MAX_COLS; j += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, j, 1, j }, true);
      }

      msxll::XLREF12 outRef;

      Assert::IsFalse(localAddressToXlRef(outRef, L"100:AB"));
      Assert::IsFalse(localAddressToXlRef(outRef, L"BuzZ100"));
      Assert::IsFalse(localAddressToXlRef(outRef, L"A1-B2"));

      Assert::IsTrue(localAddressToXlRef(outRef, L"$A1"));
      Assert::IsTrue(localAddressToXlRef(outRef, L"$A1:B$1"));
    }
  };
}