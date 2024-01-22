#include <xloil/AutoBind.h>

namespace AnotherCLib
{
  double AddNumbers(double x, double y, double z = 1)
  {
    return x + y + z;
  }

  std::vector<double> AddVectors(const std::vector<double>& x, const std::vector<double>& y)
  {
    std::vector<double> result;

    for (auto i = 0; i < std::min(x.size(), y.size()); ++i)
      result.push_back(x[i] + y[i]);
    
    return result;
  }

  std::wstring DoStringThing(const wchar_t* str)
  {
    return std::wstring(str);
  }
}

using namespace xloil::AutoBind;

using AnotherCLib::AddNumbers;
XLO_BIND(AddNumbers, 3, XLO_DEFAULTS(None, None, 1.0))
  .help(L"Example auto bind")
  .arg(L"x", L"first arg")
  .arg(L"y", L"next arg")
  .arg(L"z", L"last arg");

using AnotherCLib::AddVectors;
XLO_BIND(AddVectors, 2)
  .help(L"Example auto bind")
  .arg(L"x", L"first arg")
  .arg(L"y", L"next arg");

using AnotherCLib::DoStringThing;
XLO_BIND(DoStringThing, 1)
  .help(L"Example auto bind")
  .arg(L"str", L"first arg");
