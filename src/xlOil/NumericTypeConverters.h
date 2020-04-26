#include "TypeConverters.h"
#include <xlOilHelpers/StringUtils.h>
#include <xlOil/Throw.h>

namespace xloil
{
  struct ToDouble : public FromExcelBase<double, ToDouble>
  {
    double fromInt(int x) const { return double(x); }
    double fromBool(bool x) const { return double(x); }
    double fromDouble(double x) const { return x; }
    double fromEmpty(const double* defaultVal) const { return defaultVal ? *defaultVal : 0.0; }
    double fromError(CellError err) const
    {
      using namespace msxll;
      if (0 != ((int)err & (xlerrNull | xlerrDiv0 | xlerrNum | xlerrNA)))
        return std::numeric_limits<double>::quiet_NaN();
      XLO_THROW("Could not convert error to double");
    }
  };
  struct ToInt : public FromExcelBase<int, ToInt>
  {
    int fromInt(int x) const { return x; }
    int fromBool(bool x) const { return int(x); }
    int fromDouble(double x) const 
    {
      int i;
      if (floatingToInt(x, i))
        return i;
      XLO_THROW("Could not convert: number not an exact integer");
    }
    int fromEmpty(const int* defaultVal) const { return defaultVal ? *defaultVal : 0; }
  };

  /// Converts to bool using Excel's standard coercions for numeric types (x != 0)
  struct ToBool : public FromExcelBase<bool, ToBool>
  {
    bool fromInt(int x) const { return x != 0; }
    bool fromBool(bool x) const { return x; }
    bool fromDouble(double x) const { return x != 0.0; }
    bool fromEmpty(const bool* defaultVal) const { return defaultVal ? *defaultVal : false; }
  };
}