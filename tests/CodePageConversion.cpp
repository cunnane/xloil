#include "CppUnitTest.h"
#include <xlOil/StringUtils.h>
#include <locale>
#include <codecvt>

using std::string;
using std::wstring;
using std::codecvt_utf8;
using std::codecvt_utf16;
using std::u32string;
using xloil::utf16ToUtf8;

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace
{
  //string To_UTF8(const std::u32string &s)
  //{
  //  std::wstring_convert<std::codecvt_utf8<char32_t>, char32_t> conv;
  //  return conv.to_bytes(s);
  //}

  wstring To_UTF16(const string &s)
  {
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>, wchar_t> conv;
    return conv.from_bytes(s);
  }

  wstring To_UTF16(const std::u32string &s)
  {
    std::wstring_convert<codecvt_utf16<int32_t, 0x10ffff, std::little_endian>, int32_t> conv;
    auto p = reinterpret_cast<const int32_t *>(s.data());
    string bytes = conv.to_bytes(p, p + s.length());
    return wstring(reinterpret_cast<const wchar_t*>(bytes.c_str()), bytes.length() / sizeof(wchar_t));
  }

  //std::u32string To_UTF32(const string &s)
  //{
  //  std::wstring_convert<codecvt_utf8<char32_t>, char32_t> conv;
  //  return conv.from_bytes(s);
  //}

  //u32string To_UTF32(const std::u16string &s)
  //{
  //  const char16_t *pData = s.c_str();
  //  std::wstring_convert<codecvt_utf16<char32_t>, char32_t> conv;
  //  return conv.from_bytes(reinterpret_cast<const char*>(pData), reinterpret_cast<const char*>(pData + s.length()));
  //}

  //u32string To_UTF32(const wstring &s)
  //{
  //  auto *pData = s.c_str();
  //  std::wstring_convert<codecvt_utf16<char32_t>, char32_t> conv;
  //  return conv.from_bytes(reinterpret_cast<const char*>(pData), reinterpret_cast<const char*>(pData + s.length()));
  //}

  // MSVC bug, see: https://stackoverflow.com/questions/32055357/visual-studio-c-2015-stdcodecvt-with-char16-t-or-char32-t

#if _MSC_VER >= 1900

  string utf16_to_utf8(std::u16string utf16_string)
  {
    std::wstring_convert<std::codecvt_utf8_utf16<int16_t>, int16_t> convert;
    auto p = reinterpret_cast<const int16_t *>(utf16_string.data());
    return convert.to_bytes(p, p + utf16_string.size());
  }

#else

  std::string utf16_to_utf8(std::u16string utf16_string)
  {
    std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> convert;
    return convert.to_bytes(utf16_string);
  }

#endif

}
namespace Tests
{
	TEST_CLASS(StringEncoding)
	{
	public:
		
		TEST_METHOD(Utf16ToUtf32)
		{
      auto utf16Str = wstring(L"Hello \u4f60\u597d_z\u00df\u6c34");
      auto result = u32string(utf16Str.length() * 2, U'\0');
      auto nChars = xloil::ConvertUTF16ToUTF32()(
        result.data(), result.length(), utf16Str.data(), utf16Str.data() + utf16Str.length());
      result.resize(nChars);
      auto resultUtf16 = To_UTF16(result);
      Assert::AreEqual(utf16Str.c_str(), resultUtf16.c_str());
		}

    TEST_METHOD(Utf32ToUtf16)
    {
      auto utf32Str = u32string(U"\U00004f60\U0000597d_z\U000000df\U00006c34\U0001f34c");
      auto result = wstring(utf32Str.length() * 2, L'\0');
      auto nChars = xloil::ConvertUTF32ToUTF16()(
        result.data(), result.length(), utf32Str.data(), utf32Str.data() + utf32Str.length());
      result.resize(nChars);
      auto utf16Str = To_UTF16(utf32Str);
      Assert::AreEqual(utf16Str.c_str(), result.c_str());
    }

    TEST_METHOD(Utf16ToUtf8)
    {
      {
        auto source = L"xlõiƚ";
        auto utf8 = utf16ToUtf8(source);
        auto utf16 = To_UTF16(utf8);
        Assert::AreEqual(source, utf16.c_str());
      }
    }
	};
}
