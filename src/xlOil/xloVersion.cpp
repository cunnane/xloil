#include <xloil/Register/FuncRegistry.h>
#include <xloil/StaticRegister.h>
#include <xloil/ArrayBuilder.h>
#include <xloil/Preprocessor.h>
#include <xloil/Version.h>
#include <xloil/Preprocessor.h>

namespace xloil
{
  XLO_FUNC_START(xloVersion())
  {
    constexpr wchar_t* version =
      BOOST_PP_CAT(XLO_WSTR(XLOIL_MAJOR_VERSION),
        BOOST_PP_CAT(".", XLO_STR(XLOIL_MINOR_VERSION)));

    constexpr wchar_t* info[2][2] = { 
      { L"Version", version },
      { L"BuildDate", XLO_WSTR(__DATE__) } 
    };

    size_t stringLen = 0;
    for (auto i = 0; i < _countof(info); ++i)
      for (auto j = 0; j < _countof(info[i]); ++j)
        stringLen += wcslen(info[i][j]);
      
    ExcelArrayBuilder builder(2, 2, stringLen);
    for (auto i = 0; i < _countof(info); ++i)
      for (auto j = 0; j < _countof(info[i]); ++j)
        builder(i, j) = info[i][j];

    return returnValue(builder.toExcelObj());
  }
  XLO_FUNC_END(xloVersion).threadsafe()
    .help(L"Version info");
}