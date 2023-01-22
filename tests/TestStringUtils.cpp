#include "CppUnitTest.h"
#include <xlOil/StringUtils.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

using namespace xloil;
using std::wstring;
using std::string;

namespace Tests
{
  TEST_CLASS(StringUtils)
  {
  public:
    template<int Radix>
    void parseRoundTrip(size_t value)
    {
      char buffer[65], buf_itoa[65];

      auto len = unsignedToString<Radix>(value, buffer, sizeof(buffer));
      auto parsed = (size_t)parseUnsigned<Radix>(buffer + 0, buffer + len);

      _ui64toa_s(value, buf_itoa, _countof(buf_itoa), Radix);
      buffer[len] = '\0';
      Assert::AreEqual<string>(buf_itoa, buffer);
      Assert::AreEqual(value, parsed);
    }

    TEST_METHOD(TestIntStringParse)
    {
      for (size_t i = 1; i < 32; ++i)
      {
        size_t value = 1ull << i;
        parseRoundTrip<2>(value);
        parseRoundTrip<7>(value);
        parseRoundTrip<10>(value);
        parseRoundTrip<16>(value);
        parseRoundTrip<32>(value);
      }
      for (size_t i = 1; i < 2000; ++i)
      {
        parseRoundTrip<2>(i);
        parseRoundTrip<7>(i);
        parseRoundTrip<10>(i);
        parseRoundTrip<16>(i);
        parseRoundTrip<32>(i);
      }
    }
  };
}