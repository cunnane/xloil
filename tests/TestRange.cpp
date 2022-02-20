#include "CppUnitTest.h"
#include <xlOil/ExcelRef.h>
#include <xlOil/ExcelObj.h>
#include <chrono>
#include <iostream>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::unique_ptr;
using std::to_string;
using std::vector;

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

    void xlRefRoundTrip(const msxll::XLREF12& ref, bool lowerCase = false, bool a1style = true)
    {
      wchar_t address[std::max(XL_FULL_ADDRESS_A1_MAX_LEN, XL_FULL_ADDRESS_RC_MAX_LEN)];
      if (a1style)
        xlrefToLocalA1(ref, address, _countof(address));
      else
        xlrefToLocalRC(ref, address, _countof(address));
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
    TEST_METHOD(TestAddressParseRC)
    {
      for (auto i = 1; i < XL_MAX_ROWS; i += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, i, 1, 1 }, false, false);
      }
      for (auto j = 1; j < XL_MAX_COLS; j += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, 1, j, j }, false, false);
      }
      for (auto j = 1; j < XL_MAX_COLS; j += 100)
      {
        xlRefRoundTrip(msxll::XLREF12{ 1, j, 1, j }, true, false);
      }

      msxll::XLREF12 outRef;

      Assert::IsFalse(localAddressToXlRef(outRef, L"R1-C1"));

      Assert::IsTrue(localAddressToXlRef(outRef, L"$R1C1"));
      Assert::IsTrue(localAddressToXlRef(outRef, L"$R1C1:R2$C1"));
    }
    TEST_METHOD(TestAddressWriteA1)
    {
      uint8_t len;
      char buf[4];
      len = writeColumnName(1 - 1, buf);
      buf[len] = '\0';
      Assert::AreEqual("A", buf);
      len = writeColumnName(27 - 1, buf);
      buf[len] = '\0';
      Assert::AreEqual("AA", buf);
      len = writeColumnName(703 - 1, buf);
      buf[len] = '\0';
      Assert::AreEqual("AAA", buf);
    }
    TEST_METHOD(TestPerformanceAddressWriteA1)
    {
      using std::chrono::high_resolution_clock;
      using std::chrono::duration_cast;
      using std::chrono::duration;
      using std::chrono::milliseconds;
      using std::to_string;

      constexpr size_t NRepeats = 1;
      constexpr size_t N = 1000;
      vector<unsigned> numbers(N, 0);

      for (auto i = 0; i < N; ++i)
      {
        char s1[16], s2[16];
        auto len = unsignedToString<10>(i, s1);
        s1[len] = '\0';
        _itoa_s(i, s2, 10);
        Assert::AreEqual(s1, s2);

        len = unsignedToString<16>(i, s1);
        s1[len] = '\0';
        _itoa_s(i, s2, 16);
        Assert::AreEqual(s1, s2);
      }

      if constexpr (NRepeats > 1)
      {
        char buffer[16];

        for (auto n : numbers)
          _itoa_s(n, buffer, 10);

        auto t1 = high_resolution_clock::now();

        for (auto k = 0; k < NRepeats; ++k)
          for (auto n : numbers)
            unsignedToString<10>(n, buffer, _countof(buffer));

        auto t2 = high_resolution_clock::now();

        for (auto k = 0; k < NRepeats; ++k)
          for (auto n : numbers)
            _itoa_s(n, buffer, 10);

        auto t3 = high_resolution_clock::now();

        /* Getting number of milliseconds as a double. */
        duration<double, std::milli> method1 = t2 - t1;
        duration<double, std::milli> method2 = t3 - t2;

        Logger::WriteMessage(("unsignedToString: " + to_string(method1.count()) + "ms\n").c_str());
        Logger::WriteMessage(("_itoa_s: " + to_string(method2.count()) + "ms\n").c_str());
      }
    }
  };
}