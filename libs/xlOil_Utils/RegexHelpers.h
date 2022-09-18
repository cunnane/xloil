#pragma once
#include <xloil/Throw.h>
#include <regex>

namespace xloil
{
  inline auto parseRegexGrammar(const std::wstring_view& name)
  {
    using namespace std::regex_constants;
    if (name.empty() || name == L"ecma") return ECMAScript;
    if (name == L"basic") return basic;
    if (name == L"extended") return extended;
    if (name == L"awk") return awk;
    if (name == L"grep") return grep;
    if (name == L"egrep") return egrep;
    XLO_THROW(L"Unknown regex grammar '{}'", name);
  }

  inline auto strView(const std::wcsub_match& matchResults)
  {
    return std::wstring_view(matchResults.first, matchResults.length());
  }
  constexpr auto _GrammerHelp =
    L"(ecma) Choice of regex syntax: ecma, basic, extended, awk, grep, egrep. See https://en.cppreference.com/w/cpp/regex/syntax_option_type";
}